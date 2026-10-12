import { BackgroundImage, Cell, Row, Workbook, Worksheet } from "xlsx-jsx";

const logo =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";

export default (
  <Workbook>
    <Worksheet name="Background">
      <BackgroundImage data={logo} />
      <Row>
        <Cell>Worksheet with a tiled background image</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
