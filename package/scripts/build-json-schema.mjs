import { mkdir, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { toJsonSchema } from '@valibot/to-json-schema'

import { CategorySchema } from '../src/schema/category.js'
import { CategorySelectionSchema } from '../src/schema/categorySelection.js'
import { FieldSchema } from '../src/schema/field.js'
import { MessagesSchema } from '../src/schema/messages.js'

const rootDirectory = dirname(dirname(fileURLToPath(import.meta.url)))

const schemas = [
	['category.json', CategorySchema],
	['categorySelection.json', CategorySelectionSchema],
	['field.json', FieldSchema],
	['messages.json', MessagesSchema],
]

const outputDirectory = join(rootDirectory, 'dist', 'schema')
await mkdir(outputDirectory, { recursive: true })

for (const [filename, schema] of schemas) {
	const jsonSchema = toJsonSchema(schema, {
		errorMode: 'ignore',
		target: 'draft-07',
	})
	const outputPath = join(outputDirectory, filename)
	await writeFile(outputPath, `${JSON.stringify(jsonSchema, null, '\t')}\n`)
}
