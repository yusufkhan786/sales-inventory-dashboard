# Deploy Checklist

1. Test index.html locally — double click, upload XLSX up to 20MB, check progress doesn't stuck at 0%
2. Verify logo visible in header/sidebar (base64 embedded)
3. Check themes: Light → Dark → Midnight (header icon)
4. Check login: only Email/Password (no Google), strength meter works
5. Check Analytics and Visuals tabs work
6. Push to GitHub:
   git init
   git add .
   git commit -m "feat: SheetSense premium production"
   git branch -M main
   git remote add origin https://github.com/<username>/sheetsense.git
   git push -u origin main
7. Enable Pages in Settings → Pages → main / (root)

If Excel upload fails on Pages (CDN blocked), SheetJS fallback will auto-try unpkg CDN. Ensure internet enabled.

For custom domain: Settings → Pages → Custom domain → add CNAME file.
