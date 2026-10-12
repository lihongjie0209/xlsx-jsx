import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Data">
      <Chart type="scatter" cell="F2" x2AxisReversed={true} y2AxisReversed={true}>
        <Series categories="Data!$A$2:$A$4" values="Data!$B$2:$B$4" />
        <Series categories="Data!$C$2:$C$4" values="Data!$D$2:$D$4" secondaryAxis={true} />
      </Chart>
      <Row>
        <Cell>Primary X</Cell><Cell>Primary Y</Cell><Cell>Secondary X</Cell><Cell>Secondary Y</Cell>
      </Row>
      <Row>
        <Cell>1</Cell><Cell>10</Cell><Cell>100</Cell><Cell>1000</Cell>
      </Row>
      <Row>
        <Cell>2</Cell><Cell>20</Cell><Cell>200</Cell><Cell>2000</Cell>
      </Row>
      <Row>
        <Cell>3</Cell><Cell>30</Cell><Cell>300</Cell><Cell>3000</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
