import { InsertLocation } from "../src/word/insertLocation";
import type { Fixture } from "../test/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Platform-divergence pair with body-insertOoxml-end.fixture.ts (PC): design
// spec's "Platform selection" — insertOoxml applies on PC/Mac, rejects on
// OfficeOnline (no client-side OOXML merge engine there, a real Word Online
// product gap, not a missing capability).
export const fixture: Fixture = {
  description: "Body.insertOoxml(End) on OfficeOnline rejects at sync()",
  platform: "OfficeOnline",
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
  apply: (context) => {
    context.document.body.insertOoxml(
      `<?xml version="1.0" encoding="UTF-8"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Imported fragment.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`,
      InsertLocation.end
    );
  },
  expectRejection: /insertOoxml/,
};
