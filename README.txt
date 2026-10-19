PDF Toolkit (Android) - By Owen
================================
Features: Images -> PDF, Merge PDF, Delete PDF pages, PDF -> Word (.docx), Word (.docx) -> PDF.
All work offline on the phone.

CONVERSION NOTES
- PDF -> Word, "Editable text": rebuilds paragraphs, headings, bold/italic, alignment and page
  breaks. Tables come out as tab-separated text and pictures inside the PDF are left out.
- PDF -> Word, "Exact look": every page becomes a picture inside the .docx (looks identical, not
  editable). Use this for scanned PDFs or complex layouts.
- Word -> PDF: reads .docx only (not old .doc). Keeps text, bold/italic/underline/colour, lists,
  tables and pictures. Fonts become Helvetica / Times / Courier; headers, footers, text boxes
  and non-Latin scripts (e.g. Hindi, Chinese) are not supported.

HOW TO GET YOUR APK (free, one time, about 10 minutes)
1. Make a free account at github.com and sign in.
2. Click "+" (top right) > "New repository". Name it pdf-toolkit, pick Private, click Create.
3. On the new repository page click "uploading an existing file".
4. Unzip this project. Open the PDFToolkitApp folder, select EVERYTHING inside it
   (including the folder named .github) and drag it into the GitHub page.
   Click "Commit changes".
5. Open the "Actions" tab. Click "Build APK" on the left, then "Run workflow" > "Run workflow".
   (If it already started by itself, just open that run.)
6. Wait about 5-8 minutes until it shows a green tick.
7. Open the finished run, scroll to "Artifacts", download "PDF-Toolkit-APK" (a zip).
8. Unzip it. app-debug.apk is your app.

INSTALL ON YOUR PHONE
1. Send app-debug.apk to your phone (Google Drive, WhatsApp, USB cable...).
2. Tap it. If Android asks, allow "Install unknown apps" for the app you opened it from.
3. If Play Protect warns, tap "Install anyway" (the app is built by you, not from the Play Store).
4. Open "PDF Toolkit" from your app drawer.

If the .github folder did not upload: on GitHub click "Add file" > "Create new file",
type  .github/workflows/build-apk.yml  as the name, and paste the contents of the
file with that name from this project.

SAVING RESULTS: after creating a PDF the phone's share sheet opens - choose
"Save to Files / Drive" or open it in a PDF app.
