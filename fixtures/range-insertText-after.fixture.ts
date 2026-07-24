import { InsertLocation } from "../src/word/insertLocation";
import type { Fixture } from "../test/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const fixture: Fixture = {
  description:
    "Range.insertText(After) splices a new sibling paragraph immediately after the range",
  seedOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>First.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002"><w:r><w:t>Third.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
  // Range isn't reachable from context.document directly — Paragraph.getRange()
  // is one of the real entry points (the other being Range.search()).
  apply: (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    range.insertText("Second.", InsertLocation.after);
  },
  resultOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>First.</w:t></w:r></w:p>
          <w:p w14:paraId="00000003"><w:r><w:t>Second.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002"><w:r><w:t>Third.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
};
