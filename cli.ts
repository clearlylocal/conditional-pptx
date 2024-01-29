import { Command } from 'cliffy/command/mod.ts'
import { type Condition, conditionalize } from './conditionalize.ts'

type Params = {
	inPath: string
	outPath: string
	include: string[]
}

export async function cli() {
	await new Command()
		.name('conditionalize-pptx')
		.version('0.1.0')
		.description('Conditional content for PPTs')
		.option('-f, --file-path <file>', 'Input file path', { required: true })
		.option('-o, --out-path <file>', 'Output file path', { required: true })
		.option('-i, --include <string>', 'Content to match in slide notes', { required: true })
		.action(async ({ filePath, outPath, include }, ..._args) => {
			await run({
				inPath: filePath,
				outPath,
				include: include.split(',').map((x) => x.trim().toLowerCase()),
			})
		})
		.parse(Deno.args)
}

async function run({ inPath, outPath, include }: Params) {
	const bytes = await Deno.readFile(inPath)

	const condition: Condition = ({ notes }) => {
		const match = notes.match(/\[([^\[\]]+)\]/u)

		if (!match) return true
		const [, magicNote] = match
		const categories = magicNote.split(',').map((x) => x.trim().toLowerCase())

		return include.some((x) => categories.includes(x))
	}

	await Deno.writeFile(outPath, await conditionalize(bytes, condition))
}
