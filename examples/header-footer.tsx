import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet
      name="Print"
      header="&CQuarterly report · Page &[Page] of &[Pages]"
      footer="&LConfidential&R&D"
      headerFooterScaleWithDoc={false}
      headerFooterAlignWithPage={false}
    >
      <Row><Cell>Header/footer output sample</Cell></Row>
    </Worksheet>
  </Workbook>
);
