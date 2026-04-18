import assert from 'node:assert/strict'
import { createWriteStream } from 'node:fs'
import { mkdtemp, rm } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'

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

const VALID_CATEGORIES = {
	'observation-category': {
		name: 'Observation Category',
		appliesTo: ['observation'],
		tags: { type: 'observation-category' },
		fields: [],
	},
}

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
