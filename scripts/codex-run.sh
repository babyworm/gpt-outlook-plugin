#!/bin/bash
# Wrapper script for calling Codex CLI from Windows/VSTO
# Usage: codex-run.sh "prompt text"
# Loads nvm environment and runs codex exec in the project directory

export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && . "$NVM_DIR/nvm.sh"

cd /mnt/c/Works/gpt_outlook_plugin 2>/dev/null || cd ~

exec codex exec -s read-only "$1" 2>/dev/null
