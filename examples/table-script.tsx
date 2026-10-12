import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Measurements">
    <Table range="A1:C3" name="MeasurementsTable">
      <TableColumn format={{script:"superscript"}} />
      <TableColumn format={{script:"subscript",numberFormat:"0.00"}} />
      <TableColumn format={{bold:true,script:"superscript"}} />
    </Table>
    <Row><Cell>Value</Cell><Cell>Ratio</Cell><Cell>Count</Cell></Row>
    <Row><Cell value={12} /><Cell value={1.25} /><Cell value={3} /></Row>
    <Row><Cell value={18} /><Cell value={0.75} /><Cell value={5} /></Row>
  </Worksheet>
</Workbook>;
