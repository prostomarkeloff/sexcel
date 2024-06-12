from aioexcel import ExcelReader
import asyncio
import aiohttp


async def from_file():
    reader = ExcelReader("example.xlsx")
    print(await reader.read_cell("A", 3))
    print(await reader.sheet_size())


async def from_http():
    url = "https://filesamples.com/samples/document/xlsx/sample3.xlsx"
    client = aiohttp.ClientSession()
    file = await (await client.get(url)).read()
    reader = ExcelReader(file)
    print(await reader.read_cell("A", 1))
    await client.close()


asyncio.run(from_http())
