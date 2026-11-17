PANEL LABEL MAKER

This ZIP includes a ready-to-run web build at the top level and editable source in the source folder.

RUN THE READY-TO-USE APP
1. Extract this ZIP.
2. Open a terminal in the extracted folder and run:
   python -m http.server 8000
   On Windows, you can use: py -m http.server 8000
3. Open http://localhost:8000 in your browser.

RUN OR EDIT THE SOURCE
1. Install Node.js 20 or newer.
2. Open a terminal in the source folder.
3. Run: npm install
4. Run: npm run dev
5. Open the local address shown in the terminal (usually http://localhost:5173).

The app processes uploaded Excel files in your browser; it does not upload them to a server.
The supplied panel-label-template.xlsx is included in both the ready-to-use build and source/public.
