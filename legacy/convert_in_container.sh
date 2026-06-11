#!/usr/bin/env bash
set -e
for f in /in/*; do
  case "$f" in
    *.doc|*.DOC|*.docx|*.DOCX)
      soffice --headless --convert-to pdf --outdir /out "$f" >/dev/null 2>&1 || echo "FAIL:$f"
      ;;
  esac
done
