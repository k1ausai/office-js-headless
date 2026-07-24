// A Flat-OPC seed with three parts: document.xml (one paragraph, with a
// realistic commentRangeStart/End + commentReference anchor — unread by this
// shim's comment primitives, included only for structural realism, same as
// a real captured fixture would contain), comments.xml (three w:comment
// elements: two top-level, one threaded reply), and commentsExtended.xml
// (the w15 threading/resolution metadata: w15:paraId matches the owning
// comment's own w:p/@w14:paraId, w15:parentParaId links a reply to its
// parent's own paraId, w15:done is the per-thread resolved flag).
//
// Comment 0 ("Jane Doe", id="0"): top-level, resolved.
// Comment 1 ("John Smith", id="1"): a reply to comment 0.
// Comment 2 ("Jane Doe", id="2"): top-level, unresolved.
export const COMMENT_SEED_OOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001">
            <w:commentRangeStart w:id="0"/>
            <w:r>
              <w:t>Commented text.</w:t>
            </w:r>
            <w:commentRangeEnd w:id="0"/>
            <w:r>
              <w:commentReference w:id="0"/>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
    <pkg:xmlData>
      <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:comment w:id="0" w:author="Jane Doe" w:date="2024-01-15T10:00:00Z" w:initials="JD">
          <w:p w14:paraId="AAAAAAAA">
            <w:r><w:t>Top-level resolved comment.</w:t></w:r>
          </w:p>
        </w:comment>
        <w:comment w:id="1" w:author="John Smith" w:date="2024-01-15T11:00:00Z" w:initials="JS">
          <w:p w14:paraId="BBBBBBBB">
            <w:r><w:t>Reply to the first comment.</w:t></w:r>
          </w:p>
        </w:comment>
        <w:comment w:id="2" w:author="Jane Doe" w:date="2024-01-16T09:00:00Z" w:initials="JD">
          <w:p w14:paraId="CCCCCCCC">
            <w:r><w:t>Second top-level, unresolved comment.</w:t></w:r>
          </w:p>
        </w:comment>
      </w:comments>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/commentsExtended.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml">
    <pkg:xmlData>
      <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
        <w15:commentEx w15:paraId="AAAAAAAA" w15:done="1"/>
        <w15:commentEx w15:paraId="BBBBBBBB" w15:done="1" w15:parentParaId="AAAAAAAA"/>
        <w15:commentEx w15:paraId="CCCCCCCC" w15:done="0"/>
      </w15:commentsEx>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;
