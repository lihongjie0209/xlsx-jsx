#!/usr/bin/env bash
set -euo pipefail

readonly REMOTE_HOST="gpu-server"
readonly REMOTE_DIR="/root/code/xlsx-jsx"
if (($# == 0)); then
  set -- --all-targets --all-features --locked
fi
printf -v cargo_args ' %q' "$@"

rsync -az \
  --exclude '/target/' \
  --exclude '/.git/' \
  --exclude '/tools/openxml-validator/bin/' \
  --exclude '/tools/openxml-validator/obj/' \
  --exclude '*.xlsx' \
  --exclude '*.docx' \
  ./ "${REMOTE_HOST}:${REMOTE_DIR}/"

ssh "${REMOTE_HOST}" \
  "cd '${REMOTE_DIR}' && export PATH=/root/.cargo/bin:\$PATH CARGO_INCREMENTAL=1 RUSTUP_DIST_SERVER=https://rsproxy.cn RUSTUP_UPDATE_ROOT=https://rsproxy.cn/rustup && cargo test${cargo_args}"
