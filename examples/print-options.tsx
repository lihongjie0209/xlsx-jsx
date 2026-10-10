import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet
      name="Print Options"
      centerHorizontally={true}
      centerVertically={true}
      printBlackAndWhite={true}
      printDraft={true}
      printHeadings={true}
      pageOrder="overThenDown"
    >
      <Row><Cell>Centered monochrome draft output with row and column headings</Cell></Row>
    </Worksheet>
  </Workbook>
);
