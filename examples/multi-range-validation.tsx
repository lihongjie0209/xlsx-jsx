import { Workbook, Worksheet, DataValidation, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Status">
      <DataValidation
        range="B2:B10 D2:D10"
        type="list"
        source={'"Open,Closed"'}
        inputTitle="Choose a status"
        inputMessage="Select Open or Closed"
      />
      <Row><Cell>First status</Cell><Cell>Open</Cell><Cell>Second status</Cell><Cell>Closed</Cell></Row>
    </Worksheet>
  </Workbook>
);
