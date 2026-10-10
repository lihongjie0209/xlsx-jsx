import { Workbook, Worksheet, Row, Cell, DefinedName } from "xlsx-jsx";

export default <Workbook>
  <DefinedName name="TaxRate" formula="0.05" />
  <Worksheet name="Data">
    <DefinedName name="LocalTotal" formula="SUM(Data!$A$1:$A$10)" />
    <Row><Cell value={100} /></Row>
  </Worksheet>
</Workbook>;
