#!/bin/sh
# Testa, gera o build e publica a pasta dist no branch gh-pages (GitHub Pages), mantendo o histórico.
set -e
cd "$(dirname "$0")/.."
npm test
npm run build
cp manifest.xml dist/manifest.xml
touch dist/.nojekyll
REMOTO=$(git remote get-url origin)
TMP=$(mktemp -d)
if git ls-remote --exit-code --heads origin gh-pages >/dev/null 2>&1; then
  git clone -q --branch gh-pages --single-branch "$REMOTO" "$TMP"
else
  git init -q -b gh-pages "$TMP"
  git -C "$TMP" remote add origin "$REMOTO"
fi
find "$TMP" -mindepth 1 -maxdepth 1 ! -name .git -exec rm -rf {} +
cp -R dist/. "$TMP"/
git -C "$TMP" add -A
if git -C "$TMP" commit -q -m "Publicação $(date '+%Y-%m-%d %H:%M')"; then
  git -C "$TMP" push -q origin gh-pages
  echo "Publicado: https://rafaelmatos81.github.io/cronograma-pptx/"
else
  echo "Nada mudou desde a última publicação."
fi
rm -rf "$TMP"
