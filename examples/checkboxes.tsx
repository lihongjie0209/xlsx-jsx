import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Checklist">
    <Row>
      <Cell value={true} checkbox={true} />
      <Cell value={false} checkbox={true} fillColor="#C6EFCE" />
    </Row>
    <Row>
      <Cell value={true} />
      <Cell>Ordinary boolean and text cells remain unchanged.</Cell>
    </Row>
  </Worksheet>
</Workbook>;
