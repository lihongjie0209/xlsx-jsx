import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet
      name="Protected"
      protection={{
        selectLockedCells: true,
        selectUnlockedCells: true,
        formatCells: true,
        sort: true,
      }}
    >
      <Row>
        <Cell locked={false}>Editable input</Cell>
        <Cell formula={"=1+1"} formulaHidden={true} />
      </Row>
    </Worksheet>
  </Workbook>
);
