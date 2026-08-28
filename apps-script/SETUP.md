# TMG secure uploader — one-time setup

The secure uploader matches the typed product name against the first worksheet in the TMG Google Sheet. Matching is exact after trimming spaces and ignoring capitalisation. It uploads the image to the root of `Dougster00/TMG-Product-Catalogue` and writes the resulting filename into the matching row's `IMAGE` column.

## 1. Create a restricted GitHub token

1. In GitHub, open **Settings → Developer settings → Personal access tokens → Fine-grained tokens**.
2. Create a token with an expiration date.
3. Set **Repository access** to **Only select repositories** and choose `TMG-Product-Catalogue`.
4. Under **Repository permissions**, give **Contents** read and write access. Do not grant other permissions.
5. Generate and copy the token. Treat it like a password.

## 2. Create the Google Apps Script project

1. Open [script.google.com](https://script.google.com) while signed into the Google account that can edit the TMG catalogue sheet.
2. Choose **New project** and name it **TMG Secure Image Admin**.
3. Replace the contents of `Code.gs` with the repository file `apps-script/Code.gs`.
4. Add a new HTML file named exactly `Index` and paste in `apps-script/Index.html`.
5. Save the project.

## 3. Store the private settings

1. Open **Project Settings** in the Apps Script editor.
2. Under **Script properties**, add:
   - `GITHUB_TOKEN` — the restricted token created above.
   - `ALLOWED_EMAIL` — the Google email address that will use the uploader.
3. Save the properties. Never paste either value into `admin.html` or any GitHub file.

## 4. Deploy privately

1. Select **Deploy → New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to yourself.
4. Set **Who has access** to **Only myself**.
5. Select **Deploy**, approve the requested Google Sheets and external-request permissions, then copy the deployed `/exec` URL.
6. Open the URL while signed into the same Google account and test it with one existing product.

After the private web-app URL is available, it can be linked from the public TMG admin page. Never use an `/dev` test URL for the finished helper.

## Safety behaviour

- No match: nothing changes.
- More than one exact match: nothing changes.
- Existing GitHub filename: nothing changes.
- Unsupported image or image larger than 10 MB: nothing changes.
- The sheet is updated only after GitHub confirms the upload.
