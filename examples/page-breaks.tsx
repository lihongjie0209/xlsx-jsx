import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Print breaks" horizontalPageBreaks={[20, 40]} verticalPageBreaks={[8]}>
    <Row><Cell>Manual page breaks preserve the print layout.</Cell></Row>
  </Worksheet>
</Workbook>;
