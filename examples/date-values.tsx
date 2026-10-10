import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Dates">
    <Row>
      <Cell date="2026-10-10" />
      <Cell date="2026-10-10T14:30:12.125" bold={true} />
    </Row>
  </Worksheet>
</Workbook>;
