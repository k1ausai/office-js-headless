// Real Word.SearchOptions also has ignorePunct/ignoreSpace/matchPrefix/
// matchSuffix/matchWholeWord/matchWildcards — confirmed unused by the
// driving consumer's actual call sites (design spec's "v1 API coverage")
// and explicitly out of scope for v1 (issue #13). Only matchCase is modeled.
export interface SearchOptions {
  matchCase?: boolean;
}
