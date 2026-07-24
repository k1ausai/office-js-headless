import JSZip from "jszip";
import { InsertLocation } from "../src/word/insertLocation";
import type { Fixture } from "../test/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const fixture: Fixture = {
  description:
    "Body.insertFileFromBase64 merges the imported .docx's styles.xml into the current document",
  seedOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Seed paragraph.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
  // insertFileFromBase64 needs a real base64-encoded zip, not a hand-typed
  // OOXML string — built here via jszip itself, same as the shim's own
  // unit tests do (see body.test.ts's buildBase64Docx helper).
  apply: async (context) => {
    const zip = new JSZip();
    zip.file(
      "word/styles.xml",
      `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/></w:style></w:styles>`
    );
    const base64Docx = await zip.generateAsync({ type: "base64" });
    context.document.body.insertFileFromBase64(base64Docx, InsertLocation.end);
  },
  // Doesn't insert body content (issue #19's documented scope) — only the
  // styles.xml merge is checked, via structuralSummary's styleNames, so
  // the document body itself is unchanged.
  resultOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Seed paragraph.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
    <pkg:xmlData>
      <w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/></w:style></w:styles>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
};
