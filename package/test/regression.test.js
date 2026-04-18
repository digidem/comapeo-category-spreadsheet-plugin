import assert from 'node:assert/strict'
import { execFile } from 'node:child_process'
import { createWriteStream } from 'node:fs'
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'

import archiver from 'archiver'

import { isParseError } from '../src/lib/errors.js'
import { Reader, Writer } from '../src/index.js'

const VALID_METADATA = {
	name: 'Regression Test Config',
	version: '1.0.0',
	builderName: 'node-test',
	builderVersion: '1.0.0',
}

const VALID_CATEGORY_SELECTION = {
	observation: ['observation-category'],
	track: [],
}

const execFileAsync = promisify(execFile)
const MESSAGES_CLI_PATH = fileURLToPath(
	new URL('../bin/comapeocat-messages.mjs', import.meta.url),
)
const VALIDATE_CLI_PATH = fileURLToPath(
	new URL('../bin/comapeocat-validate.mjs', import.meta.url),
)
const VALID_CATEGORIES = {
	'observation-category': {
		name: 'Observation Category',
		appliesTo: ['observation'],
		tags: { type: 'observation-category' },
		fields: [],
	},
}

test('comapeocat-messages uses dot-prop indexes for option label message IDs', async () => {
	await withTempDir(async (tempDir) => {
		await writeJson(path.join(tempDir, 'fields', 'status.json'), {
			tagKey: 'status',
			label: 'Status',
			type: 'selectOne',
			options: [
				{ label: 'Open', value: 'open' },
				{ label: 'Closed', value: 'closed' },
			],
		})

		const { stdout, stderr } = await execFileAsync(process.execPath, [
			MESSAGES_CLI_PATH,
			tempDir,
		])
		assert.equal(stderr, '')

		const messages = JSON.parse(stdout)
		assert.deepEqual(messages['field.status.options.0.label'], {
			description: "Label for option 'open' of field 'status'",
			message: 'Open',
		})
		assert.deepEqual(messages['field.status.options.1.label'], {
			description: "Label for option 'closed' of field 'status'",
			message: 'Closed',
		})
		assert.equal(
			'field.status.options[value="open"].label' in messages,
			false,
		)
	})
})

test('Writer.finish rejects categorySelection references to missing categories', () => {
	const writer = new Writer()
	writer.setMetadata(VALID_METADATA)
	writer.addCategory('observation-category', VALID_CATEGORIES['observation-category'])
	writer.setCategorySelection({
		observation: ['missing-category'],
		track: [],
	})

	assert.throws(() => writer.finish(), {
		name: 'CategorySelectionRefError',
		message:
			/Category "missing-category" referenced by "categorySelection\.observation" is missing\./,
	})
})

test('Writer.finish rejects categorySelection references to categories with incompatible appliesTo', () => {
	const writer = new Writer()
	writer.setMetadata(VALID_METADATA)
	writer.addCategory('track-category', {
		name: 'Track Category',
		appliesTo: ['track'],
		tags: { type: 'track-category' },
		fields: [],
	})
	writer.setCategorySelection({
		observation: ['track-category'],
		track: ['track-category'],
	})

	assert.throws(() => writer.finish(), {
		name: 'InvalidCategorySelectionError',
		message:
			/Category "track-category" in categorySelection\.observation does not include "observation" in its appliesTo array/,
	})
})

test('Reader.validate rejects invalid SVG icon payloads in archives', async () => {
	await withTempDir(async (tempDir) => {
		const archivePath = path.join(tempDir, 'invalid-icon.comapeocat')
		await createArchive(archivePath, {
			'categories.json': JSON.stringify(VALID_CATEGORIES, null, 2),
			'fields.json': JSON.stringify({}, null, 2),
			'categorySelection.json': JSON.stringify(VALID_CATEGORY_SELECTION, null, 2),
			'metadata.json': JSON.stringify(
				{
					...VALID_METADATA,
					buildDateValue: Date.now(),
				},
				null,
				2,
			),
			VERSION: '1.0',
			'icons/broken.svg': '<svg><path></svg>',
		})

		const reader = new Reader(archivePath)
		try {
			await assert.rejects(reader.validate(), {
				name: 'InvalidSvgError',
				message: /Invalid SVG content/,
			})
		} finally {
			await reader.close()
		}
	})
})

test('Validate CLI prints invalid SVG errors without a stack trace', async () => {
	await withTempDir(async (tempDir) => {
		const archivePath = path.join(tempDir, 'invalid-icon.comapeocat')
		await createArchive(archivePath, {
			'categories.json': JSON.stringify(VALID_CATEGORIES, null, 2),
			'fields.json': JSON.stringify({}, null, 2),
			'categorySelection.json': JSON.stringify(VALID_CATEGORY_SELECTION, null, 2),
			'metadata.json': JSON.stringify(
				{
					...VALID_METADATA,
					buildDateValue: Date.now(),
				},
				null,
				2,
			),
			VERSION: '1.0',
			'icons/broken.svg': '<svg><path></svg>',
		})

		await assert.rejects(
			execFileAsync(process.execPath, [VALIDATE_CLI_PATH, archivePath]),
			(err) => {
				assert.equal(err?.code, 1)
				assert.equal(err?.stdout, '')
				assert.match(err?.stderr ?? '', /^Invalid SVG content\n?$/)
				assert.doesNotMatch(err?.stderr ?? '', /InvalidSvgError|\sat\s/)
				return true
			},
		)
	})
})

test('Reader.opened and Reader.validate reject invalid translation filenames', async () => {
	await withTempDir(async (tempDir) => {
		const archivePath = path.join(tempDir, 'invalid-translation-tag.comapeocat')
		await createArchive(archivePath, {
			'categories.json': JSON.stringify(VALID_CATEGORIES, null, 2),
			'fields.json': JSON.stringify({}, null, 2),
			'categorySelection.json': JSON.stringify(VALID_CATEGORY_SELECTION, null, 2),
			'metadata.json': JSON.stringify(
				{
					...VALID_METADATA,
					buildDateValue: Date.now(),
				},
				null,
				2,
			),
			VERSION: '1.0',
			'translations/invalid-tag-.json': JSON.stringify({}, null, 2),
		})

		const openedReader = new Reader(archivePath)
		try {
			await assert.rejects(openedReader.opened(), {
				name: 'InvalidTranslationFilenameError',
				message:
					/Invalid translation filename: translations\/invalid-tag-\.json\nInvalid BCP 47 tag: invalid-tag-/,
			})
		} finally {
			await openedReader.close()
		}

		const validateReader = new Reader(archivePath)
		try {
			await assert.rejects(validateReader.validate(), {
				name: 'InvalidTranslationFilenameError',
				message:
					/Invalid translation filename: translations\/invalid-tag-\.json\nInvalid BCP 47 tag: invalid-tag-/,
			})
		} finally {
			await validateReader.close()
		}
	})
})

test('Reader invalid translation filename is treated as a parse-style CLI validation error', async () => {
	await withTempDir(async (tempDir) => {
		const archivePath = path.join(tempDir, 'invalid-translation-tag.comapeocat')
		await createArchive(archivePath, {
			'categories.json': JSON.stringify(VALID_CATEGORIES, null, 2),
			'fields.json': JSON.stringify({}, null, 2),
			'categorySelection.json': JSON.stringify(VALID_CATEGORY_SELECTION, null, 2),
			'metadata.json': JSON.stringify(
				{
					...VALID_METADATA,
					buildDateValue: Date.now(),
				},
				null,
				2,
			),
			VERSION: '1.0',
			'translations/invalid-tag-.json': JSON.stringify({}, null, 2),
		})

		const reader = new Reader(archivePath)
		try {
			await assert.rejects(reader.validate(), (err) => {
				assert.equal(err?.name, 'InvalidTranslationFilenameError')
				assert.equal(isParseError(err), true)
				assert.match(String(err?.message), /Invalid translation filename: translations\/invalid-tag-\.json/)
				return true
			})
		} finally {
			await reader.close()
		}
	})
})

async function writeJson(filePath, value) {
	await mkdir(path.dirname(filePath), { recursive: true })
	await writeFile(filePath, JSON.stringify(value, null, 2))
}

async function createArchive(archivePath, entries) {
	await new Promise((resolve, reject) => {
		const output = createWriteStream(archivePath)
		const archive = archiver('zip', { zlib: { level: 9 } })

		output.on('close', resolve)
		output.on('error', reject)
		archive.on('error', reject)
		archive.pipe(output)

		for (const [name, content] of Object.entries(entries)) {
			archive.append(content, { name })
		}

		archive.finalize().catch(reject)
	})
}

async function withTempDir(fn) {
	const tempDir = await mkdtemp(path.join(os.tmpdir(), 'comapeocat-regression-'))
	try {
		await fn(tempDir)
	} finally {
		await rm(tempDir, { recursive: true, force: true })
	}
}
