import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Zoomed report" zoom={125}>
    <Row><Cell>Screen zoom remains independent of print scaling.</Cell></Row>
  </Worksheet>
</Workbook>;
