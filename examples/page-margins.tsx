import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Print margins" pageMargins={{left: 0.5, top: 1}}>
    <Row><Cell>Margins are expressed in inches.</Cell></Row>
  </Worksheet>
</Workbook>;
