import { Workbook, Worksheet, Column, Row, Cell, Image } from "xlsx-jsx";

const logo = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";

export default <Workbook>
  <Worksheet name="Custom image cells">
    <Column min={1} max={1} widthPixels={200} />
    <Image cell="A1" fit="containCenter" data={logo} altText="Centered fitted image" />
    <Row height={45}><Cell>Custom 200×60-pixel image cell</Cell></Row>
  </Worksheet>
</Workbook>;
