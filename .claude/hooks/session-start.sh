#!/bin/bash
set -euo pipefail

# Only run in remote Claude Code environments
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

echo '{"async": true, "asyncTimeout": 300000}'

install_skill() {
  local repo="$1"
  local skill="$2"
  local dir="${CLAUDE_PROJECT_DIR}/.agents/skills/${skill}"
  if [ ! -d "$dir" ]; then
    npx --yes skills add "$repo" --skill "$skill"
  fi
}

# anthropics/skills
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

# ComposioHQ/skills
install_skill https://github.com/ComposioHQ/skills composio

# supabase/agent-skills
install_skill https://github.com/supabase/agent-skills supabase-postgres-best-practices

# cloudflare/skills
install_skill https://github.com/cloudflare/skills agents-sdk
install_skill https://github.com/cloudflare/skills sandbox-sdk
install_skill https://github.com/cloudflare/skills web-perf

# VoltAgent/skills
install_skill https://github.com/VoltAgent/skills voltagent-best-practices

# google-gemini/gemini-skills
install_skill https://github.com/google-gemini/gemini-skills gemini-api-dev
install_skill https://github.com/google-gemini/gemini-skills gemini-interactions-api
install_skill https://github.com/google-gemini/gemini-skills gemini-live-api-dev

# stripe/ai
install_skill https://github.com/stripe/ai stripe-best-practices
install_skill https://github.com/stripe/ai upgrade-stripe

# callstackincubator/agent-skills
install_skill https://github.com/callstackincubator/agent-skills github
install_skill https://github.com/callstackincubator/agent-skills upgrading-react-native

# better-auth/skills
install_skill https://github.com/better-auth/skills better-auth-best-practices
install_skill https://github.com/better-auth/skills create-auth-skill
install_skill https://github.com/better-auth/skills email-and-password-best-practices
install_skill https://github.com/better-auth/skills organization-best-practices
install_skill https://github.com/better-auth/skills two-factor-authentication-best-practices

# tinybirdco/tinybird-agent-skills
install_skill https://github.com/tinybirdco/tinybird-agent-skills tinybird-python-sdk-guidelines

# sanity-io/agent-toolkit
install_skill https://github.com/sanity-io/agent-toolkit content-modeling-best-practices
install_skill https://github.com/sanity-io/agent-toolkit content-experimentation-best-practices

# firecrawl/skills
install_skill https://github.com/firecrawl/skills firecrawl-build-interact
install_skill https://github.com/firecrawl/skills firecrawl-build-scrape
install_skill https://github.com/firecrawl/skills firecrawl-build-search

# clickhouse/agent-skills
install_skill https://github.com/clickhouse/agent-skills chdb-datastore
install_skill https://github.com/clickhouse/agent-skills chdb-sql
install_skill https://github.com/clickhouse/agent-skills clickhouse-architecture-advisor
install_skill https://github.com/clickhouse/agent-skills clickhousectl-cloud-deploy

# remotion-dev/skills
install_skill https://github.com/remotion-dev/skills remotion-best-practices

# typefully/agent-skills
install_skill https://github.com/typefully/agent-skills typefully

# trycourier/courier-skills (git clone method)
if [ ! -d "${CLAUDE_PROJECT_DIR}/.agents/skills/courier-skills" ]; then
  git clone https://github.com/trycourier/courier-skills.git /tmp/courier-skills
  cp -r /tmp/courier-skills "${CLAUDE_PROJECT_DIR}/.agents/skills/courier-skills"
  rm -rf "${CLAUDE_PROJECT_DIR}/.agents/skills/courier-skills/.git"
  rm -rf /tmp/courier-skills
fi
