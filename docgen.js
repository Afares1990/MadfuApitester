/* docgen.js
   Browser DOCX generation using window.docx (UMD).
   - Blue-styled credentials table (adjust colors at top)
   - Credit-card test block (Visa 4111 1111 1111 1111, CVV 123, Expiry 05/25, Name: test)
   - Email-style layout (Subject, Greeting, Body, Signature)
   - Clickable Dashboard and API Docs links via ExternalHyperlink
   - HTML-as-.doc fallback if docx isn't available
   Attach this script in index.html after the docx UMD script tag.
*/

(function () {
  // Adjustable styling colors (hex; no leading #)
  const CREDENTIAL_HEADER_COLOR = "0B66C2"; // header text color
  const CREDENTIAL_HEADER_FILL = "0B66C2";  // header background
  const CREDENTIAL_VALUE_FILL = "E6F0FF";   // value cell background (lighter blue)

  function hasDocx() {
    return typeof window.docx !== "undefined" && window.docx !== null;
  }

  function getField(id) {
    const el = document.getElementById(id);
    return el ? el.value || "" : "";
  }

  function buildCredentialsTable(docx) {
    const {
      Table,
      TableRow,
      TableCell,
      Paragraph,
      TextRun,
      WidthType,
      BorderStyle,
    } = docx;

    function makeCell(text, isHeader) {
      return new TableCell({
        shading: {
          type: "clear",
          color: isHeader ? "FFFFFF" : "000000",
          fill: isHeader ? CREDENTIAL_HEADER_FILL : CREDENTIAL_VALUE_FILL,
        },
        margins: { top: 100, bottom: 100, left: 100, right: 100 },
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text,
                bold: isHeader,
                color: isHeader ? "FFFFFF" : "0B1724",
              }),
            ],
          }),
        ],
      });
    }

    const apiKey = getField("apiKey");
    const appCode = getField("appCode");
    const basicAuth = getField("basicAuth");
    const adminUser = getField("adminUser");
    const adminPassword = getField("adminPassword");

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [makeCell("Credential", true), makeCell("Value", true)],
        }),
        new TableRow({
          children: [makeCell("APIKey", false), makeCell(apiKey || "<empty>", false)],
        }),
        new TableRow({
          children: [makeCell("AppCode", false), makeCell(appCode || "<empty>", false)],
        }),
        new TableRow({
          children: [makeCell("Authorization (Basic)", false), makeCell(basicAuth || "<empty>", false)],
        }),
        new TableRow({
          children: [makeCell("Dashboard Admin User", false), makeCell(adminUser || "<empty>", false)],
        }),
        new TableRow({
          children: [makeCell("Dashboard Admin Password", false), makeCell(adminPassword ? "••••••••" : "<empty>", false)],
        }),
      ],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "BBDDF8" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "BBDDF8" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "BBDDF8" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "BBDDF8" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E6F0FF" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "E6F0FF" },
      },
    });
  }

  function buildCreditCardBlock(docx) {
    const { Table, TableRow, TableCell, Paragraph, TextRun } = docx;
    const ccNumber = "Visa 4111 1111 1111 1111";
    const ccCvv = "CVV 123";
    const ccExpiry = "Expiry 05/25";
    const ccName = "Name: test";

    function ccCell(text, bold) {
      return new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text, bold })] })],
      });
    }

    return new Table({
      rows: [
        new TableRow({ children: [ccCell("Card", true), ccCell(ccNumber, false)] }),
        new TableRow({ children: [ccCell("CVV", true), ccCell(ccCvv, false)] }),
        new TableRow({ children: [ccCell("Expiry", true), ccCell(ccExpiry, false)] }),
        new TableRow({ children: [ccCell("Name", true), ccCell(ccName, false)] }),
      ],
      width: { size: 70, type: "pct" },
    });
  }

  function buildEmailDoc(docx) {
    const { Document, Paragraph, TextRun, HeadingLevel, ExternalHyperlink } = docx;

    const merchantName = getField("merchantName") || "Merchant";
    const dashboardName = getField("dashboardName") || "Madfu Vendor Dashboard";
    const dashboardUrl = getField("dashboardUrl") || "https://vendor-staging-new.madfu.com.sa/login";
    const apiDocsUrl = getField("apiDocsUrl") || "https://api.madfu.com.sa/docs";

    const subjectText = `API Flow Test Results - ${merchantName}`;
    const subject = new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text: `Subject: ${subjectText}`, bold: true })],
    });

    const greeting = new Paragraph({
      children: [new TextRun({ text: `Hello ${merchantName},`, break: 1 })],
    });

    const body1 = new Paragraph({
      children: [
        new TextRun({
          text:
            "This email contains the API flow test results and environment credentials used during this run. Use the Dashboard and API Docs links below to access the systems directly.",
        }),
      ],
    });

    const credentialsTable = buildCredentialsTable(docx);
    const ccBlock = buildCreditCardBlock(docx);

    const dashboardLink = new ExternalHyperlink({
      children: [new TextRun({ text: dashboardName, style: "Hyperlink" })],
      link: dashboardUrl,
    });
    const apiDocsLink = new ExternalHyperlink({
      children: [new TextRun({ text: "API Docs", style: "Hyperlink" })],
      link: apiDocsUrl,
    });

    const linksParagraph = new Paragraph({
      children: [dashboardLink, new TextRun({ text: "  |  " }), apiDocsLink],
    });

    const signature = new Paragraph({
      children: [
        new TextRun({ text: "Regards,", break: 1 }),
        new TextRun({ text: "Madfu API Tester", break: 1, bold: true }),
      ],
    });

    return new Document({
      sections: [
        {
          properties: {},
          children: [
            subject,
            greeting,
            body1,
            new Paragraph({ text: "" }),
            credentialsTable,
            new Paragraph({ text: "" }),
            new Paragraph({ children: [new TextRun({ text: "Credit Card Test (for sandbox):", bold: true })] }),
            ccBlock,
            new Paragraph({ text: "" }),
            linksParagraph,
            new Paragraph({ text: "" }),
            signature,
          ],
        },
      ],
    });
  }

  function downloadHtmlFallback(filename) {
    const merchantName = getField("merchantName") || "Merchant";
    const dashboardUrl = getField("dashboardUrl") || "https://vendor-staging-new.madfu.com.sa/login";
    const apiDocsUrl = getField("apiDocsUrl") || "https://api.madfu.com.sa/docs";
    const apiKey = getField("apiKey") || "<empty>";
    const appCode = getField("appCode") || "<empty>";
    const basicAuth = getField("basicAuth") || "<empty>";
    const adminUser = getField("adminUser") || "<empty>";
    const adminPassword = getField("adminPassword") ? "••••••••" : "<empty>";

    const html = `
      <!doctype html>
      <html>
        <head><meta charset="utf-8"><title>${filename}</title></head>
        <body>
          <h2>Subject: API Flow Test Results - ${merchantName}</h2>
          <p>Hello ${merchantName},</p>
          <p>This email contains the API flow test results and environment credentials used during this run.</p>
          <h3>Credentials</h3>
          <table border="1" cellpadding="6">
            <tr style="background:#${CREDENTIAL_HEADER_FILL};color:#fff"><th>Credential</th><th>Value</th></tr>
            <tr><td>APIKey</td><td>${apiKey}</td></tr>
            <tr><td>AppCode</td><td>${appCode}</td></tr>
            <tr><td>Authorization (Basic)</td><td>${basicAuth}</td></tr>
            <tr><td>Dashboard Admin User</td><td>${adminUser}</td></tr>
            <tr><td>Dashboard Admin Pass</td><td>${adminPassword}</td></tr>
          </table>
          <h3>Credit Card Test</h3>
          <pre>Card: Visa 4111 1111 1111 1111
CVV: 123
Expiry: 05/25
Name: test</pre>
          <p><a href="${dashboardUrl}">Dashboard</a> | <a href="${apiDocsUrl}">API Docs</a></p>
          <p>Regards,<br/>Madfu API Tester</p>
        </body>
      </html>
    `;

    const blob = new Blob([html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = filename.replace(/[^a-z0-9_.-]/gi, "_") + ".doc";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  async function generateDocxAndDownload() {
    const filenameBase = (getField("merchantName") || "madfu-api-test").replace(/\s+/g, "_");
    const filename = `${filenameBase}_api-flow.docx`;

    if (!hasDocx()) {
      console.warn("docx UMD not found; using HTML fallback.");
      downloadHtmlFallback(filenameBase);
      return;
    }

    try {
      const { Packer } = window.docx;
      const doc = buildEmailDoc(window.docx);

      const blob = await Packer.toBlob(doc);
      if (typeof saveAs === "function") {
        saveAs(blob, filename);
      } else {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }
    } catch (err) {
      console.error("docx generation failed, falling back to HTML .doc", err);
      downloadHtmlFallback(filenameBase);
    }
  }

  function attachDocGenHandlers() {
    const btn = document.getElementById("previewGenerate");
    if (!btn) return;
    btn.addEventListener("click", function (e) {
      e.preventDefault();
      btn.disabled = false;
      btn.classList.remove("disabled");
      generateDocxAndDownload();
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", attachDocGenHandlers);
  } else {
    attachDocGenHandlers();
  }

  window.MadfuDocGen = {
    generate: generateDocxAndDownload,
    attach: attachDocGenHandlers,
  };
})();
