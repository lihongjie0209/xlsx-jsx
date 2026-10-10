import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="A4 Landscape" orientation="landscape" paperSize={9}>
      <Row><Cell>Printed in landscape on A4 paper</Cell></Row>
    </Worksheet>
  </Workbook>
);
