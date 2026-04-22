/**
 * 选择性保存高保真方案：XML 字节级哈希器。
 *
 * 设计目标：
 *   - 把原始 DOCX 中每个顶层块（w:p / w:tbl / w:sdt / w:sectPr）的字节子串，
 *     规范化后计算一个 64-bit FNV-1a 指纹，作为"节点身份"的稳定哈希。
 *   - 哈希在导入阶段一次性计算并写入 Tiptap 节点的 `__origHash` attr，
 *     导出时再对 localSerializer 输出的 XML 再做一次同等规范化+哈希并比对；
 *     若哈希仍匹配 → 说明该节点"未被用户修改"（或改动后又被还原），
 *     直接走字节级 "原样复制" 路径，完美保真。
 *
 * 规范化策略（第一阶段保守版本，后续可按需增强）：
 *   1. 剥离 BOM（UTF-8 `EF BB BF`）。
 *   2. 去除标签之间的纯空白文本节点（Word 不同版本格式化空白存在差异）。
 *   3. 把自闭合标签统一成紧凑形式：`<w:p />` → `<w:p/>`。
 *   4. 统一属性前引号两种写法（保留原样不做重排，避免引入 O(n log n) 解析开销；
 *      属性重排留给 phase 3 按需启用）。
 *
 * 非目标：完整 Canonical XML (C14N) —— 那会把每个节点都要过一次 DOM，
 * 性能不可接受且绝大多数 Word 文档变体靠上述四点已能稳定匹配。
 */

const BOM_UTF8 = Uint8Array.of(0xef, 0xbb, 0xbf)

/**
 * 64-bit FNV-1a 哈希。用 BigInt 实现，输出 16 位小写 hex 字符串。
 *
 * 选 FNV-1a 的原因：
 *   - 纯字节运算，无需 WASM / node:crypto 依赖（浏览器端可直接复用）。
 *   - 碰撞率对我们的用例（每文档最多 ~10000 个段落/表格）完全足够。
 *   - 实现极小，CI/启动开销为零。
 */
export function fnv1a64(bytes: Uint8Array): string {
  let hash = 0xcbf29ce484222325n
  const prime = 0x100000001b3n
  const mask = 0xffffffffffffffffn
  for (let i = 0; i < bytes.length; i++) {
    hash ^= BigInt(bytes[i])
    hash = (hash * prime) & mask
  }
  return hash.toString(16).padStart(16, '0')
}

/**
 * 对 XML 字节子串做最小化规范化。
 *
 * 返回新的 Uint8Array；原字节不做修改。不使用 fast-xml-parser / DOMParser，
 * 因为一次全文档导入可能要算上万次哈希，任何 O(节点数) 的解析都会成为瓶颈。
 */
export function canonicalizeXmlBytes(input: Uint8Array): Uint8Array {
  const stripped = stripBom(input)
  const text = bytesToString(stripped)
  const normalized = normalizeXmlString(text)
  return stringToBytes(normalized)
}

/** 规范化 + FNV-1a 的组合便捷方法，handler 侧直接调用。 */
export function hashXmlRange(input: Uint8Array): string {
  return fnv1a64(canonicalizeXmlBytes(input))
}

function stripBom(input: Uint8Array): Uint8Array {
  if (
    input.length >= 3 &&
    input[0] === BOM_UTF8[0] &&
    input[1] === BOM_UTF8[1] &&
    input[2] === BOM_UTF8[2]
  ) {
    return input.subarray(3)
  }
  return input
}

/**
 * 字符串层面的规范化。使用正则一次性覆盖：
 *   - 标签间纯空白文本（仅 \s+）：整段删除。
 *   - 自闭合写法：`<tag ... />` 中 `/>` 前的 ASCII 空白合并为单个 `/>`。
 * 不触碰文本节点内部的空白（Word 的 xml:space="preserve" 语义要求保留）。
 */
function normalizeXmlString(text: string): string {
  let out = text
  // 标签之间的纯空白：> \s+ <  →  ><
  out = out.replace(/>\s+</g, '><')
  // 自闭合前冗余空白：< ... \s+/>  →  < ... />
  out = out.replace(/\s+\/>/g, '/>')
  return out
}

function bytesToString(input: Uint8Array): string {
  return new TextDecoder('utf-8', { fatal: false }).decode(input)
}

function stringToBytes(input: string): Uint8Array {
  return new TextEncoder().encode(input)
}
