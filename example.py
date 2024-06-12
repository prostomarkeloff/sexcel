from aioexcel import ExcelReader
import asyncio


async def main():
    reader = ExcelReader("example.xlsx")
    print(await reader.read_cell("A", 3))
    print(await reader.sheet_size())


asyncio.run(main())
