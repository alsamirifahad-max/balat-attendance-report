#!/usr/bin/env bash
# Security policy checks for the published site. Fails the build on any violation.
set -euo pipefail
cd "$(dirname "$0")/.."
fail(){ echo "::error::$*"; exit 1; }

# 1) integrity manifest matches every shipped file
( cd site && sha256sum --quiet -c SHA256SUMS ) || fail "site/ files do not match SHA256SUMS"
listed=$(cut -d' ' -f3- site/SHA256SUMS | sort); actual=$(cd site && find . -type f ! -name SHA256SUMS | sed 's#^\./##' | sort)
[ "$listed" = "$actual" ] || fail "site/ contains files that are not in SHA256SUMS"

# 2) no inline scripts, no inline event handlers, no javascript: URLs
for f in site/*.html; do
  grep -Eqi '<script>|<script[^>]*>[^<]+</script>' "$f" && fail "inline <script> in $f"
  grep -Eqi '\son[a-z]+=' "$f" && fail "inline event handler in $f"
  grep -Eqi 'javascript:' "$f" && fail "javascript: URL in $f"
done

# 3) no third-party requests: only same-origin resources may be referenced
if grep -Eo '(src|href)="https?://[^"]+"' site/*.html; then fail "external resource referenced in HTML"; fi
if grep -Eo 'url\(["'"'"']?https?://' site/assets/*.css; then fail "external url() in CSS"; fi
if grep -Eo "fetch\(['\"]https?://" site/assets/*.js; then fail "external fetch in app code"; fi

# 4) Content-Security-Policy must stay strict
csp=$(grep -E '^\s*Content-Security-Policy\s*=' netlify.toml || true)
[ -n "$csp" ] || fail "Content-Security-Policy header missing from netlify.toml"
echo "$csp" | grep -q "default-src 'none'"        || fail "CSP must use default-src 'none'"
echo "$csp" | grep -Eq "script-src 'self'( 'wasm-unsafe-eval')?;" || fail "CSP script-src must be 'self' (plus 'wasm-unsafe-eval' for the on-device OCR)"
echo "$csp" | grep -q "frame-ancestors 'none'"     || fail "CSP must forbid framing"
echo "$csp" | sed "s/'wasm-unsafe-eval'//" | grep -Eq "unsafe-eval|script-src[^;]*unsafe-inline|\*" && fail "CSP contains unsafe-eval, script unsafe-inline or a wildcard"
for h in Strict-Transport-Security X-Content-Type-Options X-Frame-Options Referrer-Policy Permissions-Policy Cross-Origin-Opener-Policy Cross-Origin-Embedder-Policy Cross-Origin-Resource-Policy; do
  grep -Eq "^\s*$h\s*=" netlify.toml || fail "header $h missing from netlify.toml"
done
echo "All security checks passed."
