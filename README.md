# interpret-travelvu-export
Converts Travelvu-Export into custom Excel format

# Installation
This tool is based on Node.JS and tested using Node 22.
The easiest way might be using nvm:
https://github.com/nvm-sh/nvm?tab=readme-ov-file#installing-and-updating

If you are using nvm:
`
nvm use
`

Install node packages:
`
npm i
`

# Usage
`npx ts-node interpret-excel.ts <path to travelvu-export>`

This will create an excel-export named as the input file, prefixed by "-result"