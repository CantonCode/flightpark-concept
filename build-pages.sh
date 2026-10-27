#!/bin/sh
# Cloudflare Pages build: publish only the website files (not apps-script/ or repo files).
set -e
rm -rf dist && mkdir -p dist
cp -R index.html assets favicon.ico site.webmanifest robots.txt sitemap.xml _redirects dist/
echo "Built dist/ with: $(ls dist | tr '\n' ' ')"
