#!/bin/bash
set -euo pipefail

echo "=== Installing Claude Code Skills Globally ==="
echo ""

SKILLS_DIR="$HOME/.claude/skills"
mkdir -p "$SKILLS_DIR"

install_skill() {
  local repo="$1"
  local skill="$2"
  if [ -d "$SKILLS_DIR/$skill" ]; then
    echo "  ✓ $skill (already installed)"
    return 0
  fi
  echo "  → Installing $skill..."
  npx --yes skills add "$repo" --skill "$skill" --global 2>/dev/null || \
  npx --yes skills add "$repo" --skill "$skill" 2>/dev/null || \
  echo "  ✗ Failed: $skill"
}

install_git_skill() {
  local repo="$1"
  local name="$2"
  if [ -d "$SKILLS_DIR/$name" ]; then
    echo "  ✓ $name (already installed)"
    return 0
  fi
  echo "  → Installing $name..."
  local tmp=$(mktemp -d)
  git clone --quiet "$repo" "$tmp" 2>/dev/null
  cp -r "$tmp" "$SKILLS_DIR/$name"
  rm -rf "$SKILLS_DIR/$name/.git"
  rm -rf "$tmp"
}

echo "[1/13] Anthropic Skills..."
install_skill https://github.com/anthropics/skills theme-factory
install_skill https://github.com/anthropics/skills brand-guidelines
install_skill https://github.com/anthropics/skills skill-creator
install_skill https://github.com/anthropics/skills internal-comms
install_skill https://github.com/anthropics/skills pptx
install_skill https://github.com/anthropics/skills xlsx
install_skill https://github.com/anthropics/skills pdf
install_skill https://github.com/anthropics/skills algorithmic-art
install_skill https://github.com/anthropics/skills canvas-design
install_skill https://github.com/anthropics/skills frontend-design
install_skill https://github.com/anthropics/skills slack-gif-creator

echo "[2/13] Composio..."
install_skill https://github.com/ComposioHQ/skills composio

echo "[3/13] Supabase Postgres..."
install_skill https://github.com/supabase/agent-skills supabase-postgres-best-practices

echo "[4/13] Cloudflare..."
install_skill https://github.com/cloudflare/skills agents-sdk
install_skill https://github.com/cloudflare/skills sandbox-sdk
install_skill https://github.com/cloudflare/skills web-perf

echo "[5/13] VoltAgent..."
install_skill https://github.com/VoltAgent/skills voltagent-best-practices

echo "[6/13] Google Gemini..."
install_skill https://github.com/google-gemini/gemini-skills gemini-api-dev
install_skill https://github.com/google-gemini/gemini-skills gemini-interactions-api
install_skill https://github.com/google-gemini/gemini-skills gemini-live-api-dev

echo "[7/13] Stripe..."
install_skill https://github.com/stripe/ai stripe-best-practices
install_skill https://github.com/stripe/ai upgrade-stripe

echo "[8/13] Callstack (GitHub & React Native)..."
install_skill https://github.com/callstackincubator/agent-skills github
install_skill https://github.com/callstackincubator/agent-skills upgrading-react-native

echo "[9/13] Better Auth..."
install_skill https://github.com/better-auth/skills better-auth-best-practices
install_skill https://github.com/better-auth/skills create-auth-skill
install_skill https://github.com/better-auth/skills email-and-password-best-practices
install_skill https://github.com/better-auth/skills organization-best-practices
install_skill https://github.com/better-auth/skills two-factor-authentication-best-practices

echo "[10/13] Tinybird..."
install_skill https://github.com/tinybirdco/tinybird-agent-skills tinybird-python-sdk-guidelines

echo "[11/13] Sanity..."
install_skill https://github.com/sanity-io/agent-toolkit content-modeling-best-practices
install_skill https://github.com/sanity-io/agent-toolkit content-experimentation-best-practices

echo "[12/13] Firecrawl..."
install_skill https://github.com/firecrawl/skills firecrawl-build-interact
install_skill https://github.com/firecrawl/skills firecrawl-build-scrape
install_skill https://github.com/firecrawl/skills firecrawl-build-search

echo "[13/13] ClickHouse & Remotion & Typefully & Courier..."
install_skill https://github.com/clickhouse/agent-skills chdb-datastore
install_skill https://github.com/clickhouse/agent-skills chdb-sql
install_skill https://github.com/clickhouse/agent-skills clickhouse-architecture-advisor
install_skill https://github.com/clickhouse/agent-skills clickhousectl-cloud-deploy
install_skill https://github.com/remotion-dev/skills remotion-best-practices
install_skill https://github.com/typefully/agent-skills typefully
install_git_skill https://github.com/trycourier/courier-skills.git courier-skills

echo ""
echo "=== Done! ==="
echo "Installed 42 skills to: $SKILLS_DIR"
echo "They are now available in ALL your Claude Code projects."
