import { Workbook, Worksheet, Row, Cell, TextRun } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Rich text">
    <Row>
      <Cell horizontalAlign="left">
        <TextRun>Status: </TextRun>
        <TextRun bold={true} fontColor="#008000">Ready</TextRun>
        <TextRun italic={true}> — verified</TextRun>
        <TextRun underline="double" strikethrough={true} script="superscript" fontName="Aptos" fontSize={12.5} fontFamily={3} fontCharset={128}> *</TextRun>
      </Cell>
    </Row>
  </Worksheet>
</Workbook>;
