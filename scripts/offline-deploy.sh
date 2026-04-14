#!/bin/bash
# Offline deployment script for collabedit-doc-converter
# Run on an internet-connected machine first, then transfer to offline machine.

set -e

IMAGES_FILE="doc-converter-images.tar"

case "${1:-}" in
  save)
    echo "=== Pulling and building images ==="
    docker pull libreofficedocker/libreoffice-unoserver:alpine3.22
    docker compose build converter

    echo "=== Saving images to ${IMAGES_FILE} ==="
    docker save \
      libreofficedocker/libreoffice-unoserver:alpine3.22 \
      collabedit-doc-converter-converter:latest \
      -o "${IMAGES_FILE}"

    echo "=== Done. Transfer ${IMAGES_FILE}, docker-compose.yml, .env, and fonts/ to the offline machine ==="
    ;;

  load)
    echo "=== Loading images from ${IMAGES_FILE} ==="
    docker load -i "${IMAGES_FILE}"

    echo "=== Starting services ==="
    docker compose up -d

    echo "=== Done. Check with: docker compose ps ==="
    ;;

  *)
    echo "Usage: $0 {save|load}"
    echo "  save  - Build images and save to tar (run on internet-connected machine)"
    echo "  load  - Load images from tar and start (run on offline machine)"
    exit 1
    ;;
esac
