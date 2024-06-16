from aioexcel import ExcelReader
import asyncio
import aiohttp


async def from_file():
    reader = ExcelReader("../sample.xlsx")
    # before accessing the values we have to read an excel file into memory
    await reader.read_into_memory()
    print(reader.read_cell("A", 3))
    print(reader.sheet_size())

    # E2 is a formula (=A1+A5)
    # without `calculate` it won't be evaluated
    print(reader.read_cell("E", 2))
    # with it - it will be
    print(reader.read_cell("E", 2, calculate=True))


async def from_http():
    url = "https://filesamples.com/samples/document/xlsx/sample3.xlsx"
    client = aiohttp.ClientSession()
    file = await (await client.get(url)).read()
    await client.close()
    reader = ExcelReader(file)
    await reader.read_into_memory()
    print(reader.read_cell("D", 4))


asyncio.run(from_file())
