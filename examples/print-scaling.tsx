import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Print scale" printScale={80}>
    <Row><Cell>Print at 80 percent.</Cell></Row>
  </Worksheet>
  <Worksheet name="Fit to pages" fitToPages={{width: 1, height: 0}}>
    <Row><Cell>Fit one page wide and let the height span as needed.</Cell></Row>
  </Worksheet>
</Workbook>;
