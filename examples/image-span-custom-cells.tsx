import { Workbook, Worksheet, Column, Image, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Mixed image geometry">
    <Column min={1} max={1} widthPixels={100} />
    <Column min={3} max={3} widthPixels={120} />
    <Image
      cell="A1"
      width={190}
      height={80}
      offsetX={10}
      offsetY={5}
      data="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg=="
    />
    <Row height={22.5}><Cell>Custom height</Cell></Row>
    <Row><Cell>Default height</Cell></Row>
    <Row height={30}><Cell>Custom height</Cell></Row>
  </Worksheet>
</Workbook>;
