import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Formula rules">
      <ConditionalFormat
        range="A2:A11"
        type="formula"
        formula1="MOD(A2,2)=0"
        fillColor="#DDEBF7"
        italic={true}
        stopIfTrue={true}
      />
      <Row>
        <Cell>Value</Cell>
      </Row>
      {Array.from({ length: 10 }, (_, index) => (
        <Row>
          <Cell>{index + 1}</Cell>
        </Row>
      ))}
    </Worksheet>
  </Workbook>
);
