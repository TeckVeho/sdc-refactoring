import { JmmEntryClient } from "../jmm-entry-client";

export default function Jmm60Page() {
  return (
    <JmmEntryClient variant="60" apiPath="/api/dosimetry/jmm60" caption="JMM60φ 記入用紙" />
  );
}
