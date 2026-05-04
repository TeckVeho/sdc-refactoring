import { JmmEntryClient } from "../jmm-entry-client";

export default function Jmm90Page() {
  return (
    <JmmEntryClient variant="90" apiPath="/api/dosimetry/jmm90" caption="JMM90φ 記入用紙" />
  );
}
