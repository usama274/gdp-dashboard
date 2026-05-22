#!/usr/bin/env bash
# Deploy the latest GHCR image to Render using the Render API.
# Usage:
#   RENDER_API_KEY=... RENDER_SERVICE_ID=... ./render-deploy.sh
set -e

if [ -z "$RENDER_API_KEY" ] || [ -z "$RENDER_SERVICE_ID" ]; then
  echo "ERROR: RENDER_API_KEY and RENDER_SERVICE_ID must be set."
  echo "Usage: RENDER_API_KEY=... RENDER_SERVICE_ID=... ./render-deploy.sh"
  exit 1
fi

IMAGE="ghcr.io/${GITHUB_REPOSITORY_OWNER:-$(git config --get remote.origin.url | sed -E 's#.*[:/](.*)/.*#\1#')}/gdp-dashboard:latest"

echo "Deploying $IMAGE to Render service $RENDER_SERVICE_ID"

curl -X PATCH "https://api.render.com/v1/services/$RENDER_SERVICE_ID" \
  -H "Authorization: Bearer $RENDER_API_KEY" \
  -H "Content-Type: application/json" \
  -d "{\"dockerImage\":\"$IMAGE\"}"

echo "Render deployment request sent."
