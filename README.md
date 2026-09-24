# Girls Inc. Outputs and Outcomes

A single-page report for Girls Inc. of Columbus & Phenix-Russell showing clients served and outcome metrics. It is one self-contained HTML file with no build step, and it follows the Girls Inc. brand palette:

| Color | Hex |
|---|---|
| Girls Inc Red | `#ED1849` |
| Girls Inc Grey | `#949CA1` |
| Black | `#000000` |
| White | `#FFFFFF` |

## View locally

Open `index.html` in any browser, or serve the folder:

```bash
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Publish with GitHub Pages

1. Push this repo to GitHub.
2. Go to **Settings → Pages** and set **Source** to **GitHub Actions**.
3. Push to `main`. The workflow in `.github/workflows/pages.yml` deploys the site automatically.

## Editing

- Colors are CSS variables at the top of the `<style>` block in `index.html` (`--accent`, `--muted`, `--line`, `--track`, and so on), with separate light and dark values.
- Fonts (Bricolage Grotesque, Figtree) load from Google Fonts with system fallbacks.
