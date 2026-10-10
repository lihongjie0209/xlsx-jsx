import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Cell borders">
    <Row>
      <Cell borderTop="thin" borderTopColor="#4472C4" borderLeft="double"
            borderDiagonal="thin" borderDiagonalColor="#70AD47" borderDiagonalType="upDown">
        Header with blue top, double left, and green diagonal edges
      </Cell>
      <Cell value={42} borderRight="dotted" borderRightColor="#C00000" />
      <Cell formula="=1+1" borderBottom="dashed" numberFormat="0.00" />
    </Row>
  </Worksheet>
</Workbook>;
