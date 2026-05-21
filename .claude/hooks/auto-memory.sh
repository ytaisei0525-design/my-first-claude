#!/bin/bash
# Auto-memory フック: セッション終了時に重要な情報を memory.md へ自動保存する

MEMORY_FILE="memory.md"
TIMESTAMP=$(date "+%Y-%m-%d %H:%M")

# stdinからフック入力を読み取る
INPUT=$(cat)

# トランスクリプトパスを抽出
TRANSCRIPT_PATH=$(echo "$INPUT" | python3 -c "
import sys, json
try:
    data = json.load(sys.stdin)
    print(data.get('transcript_path', ''))
except Exception:
    print('')
" 2>/dev/null)

# memory.md が存在しない場合は初期化
if [ ! -f "$MEMORY_FILE" ]; then
    cat > "$MEMORY_FILE" << 'INIT'
# プロジェクトメモリ

このファイルは Claude Code の Auto Memory フックによって自動管理されます。
重要な決定事項、学習内容、プロジェクト固有の情報がここに蓄積されます。

---

INIT
fi

# トランスクリプトから要約を抽出して保存
if [ -n "$TRANSCRIPT_PATH" ] && [ -f "$TRANSCRIPT_PATH" ]; then
    SUMMARY=$(python3 << PYEOF
import json
import sys

try:
    with open("$TRANSCRIPT_PATH", "r", encoding="utf-8") as f:
        transcript = json.load(f)

    # アシスタントの最後のメッセージを抽出
    messages = transcript if isinstance(transcript, list) else transcript.get("messages", [])
    assistant_msgs = [
        m for m in messages
        if isinstance(m, dict) and m.get("role") == "assistant"
    ]

    if not assistant_msgs:
        sys.exit(0)

    last_msg = assistant_msgs[-1]
    content = last_msg.get("content", "")
    if isinstance(content, list):
        text_parts = [c.get("text", "") for c in content if isinstance(c, dict) and c.get("type") == "text"]
        content = " ".join(text_parts)

    # 最初の200文字を要約として使用
    summary = content[:200].replace("\n", " ").strip()
    if len(content) > 200:
        summary += "..."

    print(summary)
except Exception as e:
    pass
PYEOF
)

    if [ -n "$SUMMARY" ]; then
        cat >> "$MEMORY_FILE" << ENTRY
## セッション記録 - $TIMESTAMP

**要約:** $SUMMARY

---

ENTRY
    else
        # 要約が取得できなかった場合はタイムスタンプのみ記録
        cat >> "$MEMORY_FILE" << ENTRY
## セッション記録 - $TIMESTAMP

*(このセッションの詳細は記録されませんでした)*

---

ENTRY
    fi
else
    # トランスクリプトがない場合はタイムスタンプのみ記録
    cat >> "$MEMORY_FILE" << ENTRY
## セッション記録 - $TIMESTAMP

*(トランスクリプトが見つかりませんでした)*

---

ENTRY
fi

echo "Auto-memory: memory.md を更新しました ($TIMESTAMP)"
