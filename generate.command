#!/usr/bin/env bash
# HasilKira — macOS double-click launcher.
#
# In Finder, double-clicking a ".command" file opens Terminal and runs it.
# The real work lives in generate.sh (shared with Linux); this wrapper just
# moves into this folder, runs it, and pauses so you can read the output.

cd "$(dirname "$0")" || exit 1
bash ./generate.sh
status=$?

echo
if [ "$status" -eq 0 ]; then
  echo "完成 / Selesai / Done  —  你可以关闭此窗口 / boleh tutup tetingkap ini."
else
  echo "出错了 / Ralat / Something went wrong — see the messages above."
fi
echo "（按任意键关闭 · tekan sebarang kekunci · press any key to close）"
read -n 1 -s -r
echo
