import { Workbook, Worksheet, DataValidation, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Status">
      <DataValidation
        range="B2:B10"
        type="list"
        source={'"Open,Closed"'}
        inputTitle="Choose a status"
        inputMessage="Select Open or Closed"
        errorTitle="Invalid status"
        errorMessage="Choose one of the listed values"
      />
      <Row><Cell>Task</Cell><Cell>Status</Cell></Row>
      <Row><Cell>Prepare report</Cell><Cell>Open</Cell></Row>
    </Worksheet>
  </Workbook>
);
