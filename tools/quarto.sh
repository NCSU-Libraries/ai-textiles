#!/bin/sh
# Runs Quarto 1.10.18, the version the Pages workflow pins, from Quarto's official image, so nothing needs installing.
#
#   tools/quarto.sh render site
#   tools/quarto.sh preview site      # then open http://localhost:4848
#
# Run it from the repository root, with Docker running. The image is linux/amd64 only; on an Apple chip Docker
# emulates it, and a full render still takes about 10 seconds.
image=ghcr.io/quarto-dev/quarto:1.10.18
tty=
[ -t 0 ] && tty=-it
if [ "$1" = preview ]; then
  shift
  exec docker run --rm $tty --platform linux/amd64 -p 4848:4848 -v "$PWD":/work -w /work "$image" \
    quarto preview "$@" --port 4848 --host 0.0.0.0 --no-browser
fi
exec docker run --rm --platform linux/amd64 -v "$PWD":/work -w /work "$image" quarto "$@"
