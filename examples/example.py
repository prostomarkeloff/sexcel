from aioexcel import ExcelReader
import asyncio
import aiohttp
import logging

logging.basicConfig(level="DEBUG")


async def from_file():
    reader = ExcelReader("../sample.xlsx")
    print(await reader.read_cell("A", 3))
    print(await reader.sheet_size())

    # E2 is a formula (=A1+A5)
    # without `calculate` it won't be evaluated
    print(await reader.read_cell("E", 2))
    # with it - it will be
    print(await reader.read_cell("E", 2, calculate=True))


async def from_http():
    url = "https://filesamples.com/samples/document/xlsx/sample3.xlsx"
    client = aiohttp.ClientSession()
    file = await (await client.get(url)).read()
    await client.close()
    reader = ExcelReader(file)
    print(await reader.read_cell("D", 4))


asyncio.run(from_file())
