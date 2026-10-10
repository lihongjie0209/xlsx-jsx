import { Workbook, Worksheet, Row, Cell, Column } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Outline directions" summaryBelow={false} summaryRight={false}>
    <Column min={1} max={1} outlineLevel={1} />
    <Column min={2} max={2} outlineLevel={1} />
    <Row outlineLevel={1}><Cell>Detail row</Cell></Row>
    <Row outlineLevel={1}><Cell>Summary row above details</Cell></Row>
  </Worksheet>
</Workbook>;
