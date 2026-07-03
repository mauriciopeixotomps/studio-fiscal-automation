#!/usr/bin/env bash
# Deploy automatizado: builda o dashboard + push para GitHub Pages
set -e

WORKDIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_FILE="$WORKDIR/.github_token"
PUBLISH="$WORKDIR/publish"
OWNER="mauriciopeixotomps"
REPO="Dash-studio-fiscal"
TMP_REPO="/tmp/deploy-repo-studio-fiscal"

if [ -n "$GITHUB_DEPLOY_TOKEN" ]; then
    TOKEN="$GITHUB_DEPLOY_TOKEN"
elif [ -f "$TOKEN_FILE" ]; then
    TOKEN=$(cat "$TOKEN_FILE" | tr -d '[:space:]')
else
    echo "✗ Token não encontrado (env GITHUB_DEPLOY_TOKEN ou arquivo $TOKEN_FILE)"
    exit 1
fi

# 1) Build dashboard
echo "→ Build dashboard..."
cd "$WORKDIR" && python3 build_dashboard.py

# 2) Cria o repo (se ainda não existir) e liga o GitHub Pages
STATUS=$(curl -s -o /dev/null -w "%{http_code}" -H "Authorization: token $TOKEN" \
    "https://api.github.com/repos/$OWNER/$REPO")
if [ "$STATUS" = "404" ]; then
    echo "→ Repo não existe, criando $OWNER/$REPO..."
    curl -s -H "Authorization: token $TOKEN" -H "Content-Type: application/json" \
        -d "{\"name\":\"$REPO\",\"description\":\"Dashboard de vendas - Studio Fiscal\",\"auto_init\":true}" \
        https://api.github.com/user/repos > /dev/null
    sleep 3
    echo "→ Ativando GitHub Pages (branch main, path /)..."
    curl -s -H "Authorization: token $TOKEN" -H "Content-Type: application/json" \
        -d '{"source":{"branch":"main","path":"/"}}' \
        "https://api.github.com/repos/$OWNER/$REPO/pages" > /dev/null
fi

# 3) Clone & sync
echo "→ Clonando repo..."
rm -rf "$TMP_REPO"
git clone -q "https://x-access-token:${TOKEN}@github.com/$OWNER/$REPO.git" "$TMP_REPO"

cd "$TMP_REPO"
cp "$PUBLISH/index.html" .

# 4) Commit & push (só se houver mudanças)
git -c user.name="Mauricio Peixoto" -c user.email="mauricio.peixoto.mps@gmail.com" add -A
if git diff --cached --quiet; then
    echo "= Sem mudanças. Nada a fazer."
else
    MSG="${1:-Atualiza dashboard $(date +'%d/%m/%Y %H:%M')}"
    git -c user.name="Mauricio Peixoto" -c user.email="mauricio.peixoto.mps@gmail.com" \
        commit -q -m "$MSG"
    git push -q origin main
    echo "✓ Push concluído: $MSG"
    echo "  URL: https://$OWNER.github.io/$REPO/"
fi
