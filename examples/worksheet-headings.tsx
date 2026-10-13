import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Data" showRowColHeaders={false} showGridlines={false}>
      <Row><Cell>Values without screen headings or gridlines</Cell></Row>
    </Worksheet>
  </Workbook>
);
