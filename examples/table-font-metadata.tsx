import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Fonts">
    <Table range="A1:A3" name="FontMetadataTable">
      <TableColumn format={{fontName:"Aptos & Serif",fontSize:12.5,fontFamily:3,fontCharset:128}} />
    </Table>
    <Row><Cell>Value</Cell></Row>
    <Row><Cell value={42} /></Row>
    <Row><Cell value={84} /></Row>
  </Worksheet>
</Workbook>;
