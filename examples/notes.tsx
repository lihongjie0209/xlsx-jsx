import { Workbook, Worksheet, Note, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Reviews">
    <Note
      cell="B3"
      text="Confirm the source value before publishing."
      author="Agent"
      authorPrefix={false}
      width={200}
      height={100}
      backgroundColor="#FFF2CC"
      altText="Review instructions"
      movement="moveAndSizeAfter"
    />
    <Row><Cell>Item</Cell><Cell>Value</Cell></Row>
    <Row hidden={true}><Cell>Hidden row</Cell><Cell>0</Cell></Row>
    <Row><Cell>Revenue</Cell><Cell>1200</Cell></Row>
  </Worksheet>
</Workbook>;
