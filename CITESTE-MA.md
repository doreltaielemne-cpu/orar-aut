# Orar Automatică UTCN, actualizat automat

Site-ul (de ex. `utcna.pages.dev`) se reface singur din fișierul Excel oficial al secției:

```
SharePoint UTCN (Excel oficial)
      │  GitHub Actions: la fiecare oră, luni–vineri, 7–22
      ▼
scripts/build_site.py  → citește Excelul, construiește pagina (dist/)
      │  doar dacă orarul s-a schimbat
      ▼
Cloudflare Pages (utcna.pages.dev)
```

Dacă descărcarea eșuează sau Excelul pare stricat (lipsesc foi sau grupe, ori are mult mai puține ore decât versiunea precedentă), **nu se publică nimic**. Site-ul rămâne pe ultima versiune bună, iar GitHub îți trimite un mail că rularea a eșuat.

## Instalare (o singură dată, cam 10 minute)

### 1. Repository pe GitHub
1. Fă-ți cont pe github.com, dacă nu ai deja.
2. Apasă **New repository**, dă-i un nume (de ex. `orar-aut`) și alege **Public**. La public, GitHub Actions e gratuit fără limită.
3. În repository apasă **Add file → Upload files**. Trage **tot conținutul** acestui folder, inclusiv folderul `.github`, apoi apasă **Commit changes**.
   - Dacă browserul nu urcă folderul `.github` (e ascuns pe unele sisteme), creează fișierul manual cu **Add file → Create new file**. La nume scrie `.github/workflows/update.yml` și lipește conținutul fișierului din zip.

### 2. Token Cloudflare
1. În Cloudflare mergi la **My Profile** (iconița de om, dreapta sus) → **API Tokens** → **Create Token** → **Create Custom Token**.
2. Nume: `orar-aut`. La **Permissions** alege **Account → Cloudflare Pages → Edit**.
3. Apasă **Continue to summary**, apoi **Create Token**. Copiază tokenul: îl vezi o singură dată.
4. **Account ID** e șirul lung din adresa paginii Cloudflare, de forma `dash.cloudflare.com/<ACCOUNT_ID>/...`.

### 3. Setări în GitHub
În repository mergi la **Settings → Secrets and variables → Actions**.
- Tab-ul **Secrets** → **New repository secret**:
  - `CLOUDFLARE_API_TOKEN` = tokenul de la pasul 2
  - `CLOUDFLARE_ACCOUNT_ID` = Account ID
- Tab-ul **Variables** → **New repository variable**:
  - `CF_PROJECT` = `utcna` (numele proiectului tău din Cloudflare Pages)
  - `ORAR_URL` = linkul Excelului oficial din anunțul de pe aut.utcluj.ro

### 4. Prima rulare
În tab-ul **Actions** apasă pe **Actualizează orarul**, apoi pe **Run workflow**. După un minut site-ul e publicat. De acum încolo rulează singur.

## Dacă SharePoint cere autentificare

Atunci descărcarea automată nu merge și prima rulare dă eroare („SharePoint nu a dat un fișier Excel”). Rămâne varianta semi-automată:

1. Descarci Excelul oficial (în Excel Online: File → Save as → Download a copy).
2. În GitHub deschizi `input/orar.xlsx` și apeși **Upload files** peste el. Merge și de pe telefon.
3. Site-ul se reface și se publică singur în cam un minut.

## Ce e în folder

| Fișier | Rol |
|---|---|
| `scripts/xparse.py` | citește foile Excel: celule unite, borduri, diagonale (săpt. pare/impare), coloane ascunse |
| `scripts/sitedata.py` | numele materiilor, tipul orei, sala, adresa, săptămânile |
| `scripts/build_site.py` | leagă totul și scrie site-ul în `dist/`, cu verificări de siguranță |
| `site/template.html` | designul paginii |
| `config.json` | săptămânile semestrului, zilele libere, notele afișate pe pagini |
| `input/orar.xlsx` | ultima versiune a Excelului oficial |
| `data/last.json` | datele publicate ultima dată (pentru „s-a schimbat ceva?”) |

## Semestrul următor
În `config.json` schimbi lista `weeks` (datele săptămânilor), `holidays` și `vacation`. Dacă apare un link nou la Excel, schimbi și variabila `ORAR_URL`.

## Test local
```
pip install openpyxl
python scripts/build_site.py input/orar.xlsx dist
```
Apoi deschizi `dist/index.html`.
