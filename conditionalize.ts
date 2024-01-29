import { BlobReader, BlobWriter, type Entry as _Entry, TextWriter, ZipReader, ZipWriter } from 'zipjs'
import { type Cheerio, type Element, load } from 'cheerio'
import { join, relative } from 'std/path/mod.ts'

// https://github.com/gildas-lormeau/zip.js/issues/371
type Entry = _Entry & { getData: Exclude<_Entry['getData'], undefined> }

type Slide = {
	path: string
}

type ConditionParams = {
	notes: string
}

export type Condition = (params: ConditionParams) => boolean

export async function conditionalize(fileBytes: Uint8Array, condition: Condition) {
	const blob = new Blob([fileBytes])

	const blobReader = new BlobReader(blob)
	const zipReader = new ZipReader(blobReader)

	const blobWriter = new BlobWriter()
	const zipWriter = new ZipWriter(blobWriter)

	const entries = await zipReader.getEntries() as Entry[]
	const entriesByPath = new Map(entries.map((e) => [
		e.filename,
		e,
	]))

	const presentationPath = 'ppt/presentation.xml'
	const relsPath = 'ppt/_rels/presentation.xml.rels'
	const specialPaths = [
		relsPath,
		presentationPath,
	] as const

	const excludeSlides: Slide[] = []

	function relToSlide($rel: Cheerio<Element>, pathRelativeTo: string): Slide {
		const path = join('ppt', relative(pathRelativeTo, $rel.attr('Target') ?? ''))

		const slide = { path }

		return slide
	}

	function toRelsPath(path: string) {
		return path.replace(/([^/]+)\.xml$/, '_rels/$1.xml.rels')
	}

	{
		const relsEntry = entries.find((x) => x.filename === relsPath)!
		const tw = new TextWriter()
		const content = await relsEntry.getData(tw)

		const $ = load(content, { xml: true })
		const rels = $('Relationship[Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"]')

		for (const rel of rels) {
			const $rel = $(rel)

			const slidePath = join('ppt', $rel.attr('Target') ?? '')
			const slide = relToSlide($rel, '.')

			const slideRelsPath = toRelsPath(slidePath)
			const entry = entriesByPath.get(slideRelsPath)

			if (!entry) {
				// slide is not included
				$rel.remove()
				continue
			}

			const tw = new TextWriter()
			const content = await entry.getData(tw)

			{
				const $ = load(content, { xml: true })
				const $notesSlideRrel = $(
					'Relationship[Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]',
				)

				const targetAttr = $notesSlideRrel.attr('Target')
				if (!targetAttr) continue

				const notesSlide = relToSlide($notesSlideRrel, '..')
				const notesSlideEntry = entriesByPath.get(notesSlide.path)

				if (!notesSlideEntry) {
					continue
				}

				const tw = new TextWriter()
				const notesContent = await notesSlideEntry.getData(tw)
				const notes = load(notesContent, { xml: true })('a\\:r').text().trim()

				const include = condition({ notes })

				if (!include) {
					// category not matched
					$rel.remove()
					excludeSlides.push(slide, notesSlide)
				}
			}
		}

		const edited = $.xml()
		zipWriter.add(relsEntry.filename, new Blob([edited]).stream())
	}

	const excludeSlidePaths = excludeSlides.map((x) => x.path)
	const excludeSlideSubPaths = excludeSlidePaths.map((x) => x.slice(4))

	{
		const presentationEntry = entries.find((x) => x.filename === presentationPath)!
		const tw = new TextWriter()
		const content = await presentationEntry.getData(tw)

		const excludeRelIds: string[] = []
		{
			const presentationRelsPath = toRelsPath(presentationPath)
			const presentationRelsEntry = entries.find((x) => x.filename === presentationRelsPath)!
			const tw = new TextWriter()
			const relsContent = await presentationRelsEntry.getData(tw)
			const $ = load(relsContent, { xml: true })
			for (const rel of $('Relationship')) {
				const $rel = $(rel)
				const relId = $rel.attr('Id')!
				const path = $rel.attr('Target')!
				const exclude = excludeSlideSubPaths.includes(path)
				if (exclude) excludeRelIds.push(relId)
			}
		}

		const $ = load(content, { xml: true })

		for (const x of $('p\\:sldIdLst p\\:sldId')) {
			const $x = $(x)

			if (excludeRelIds.includes($x.attr('r:id')!)) {
				$x.remove()
			}
		}

		const edited = $.xml()
		zipWriter.add(presentationEntry.filename, new Blob([edited]).stream())
	}

	for (const entry of entries) {
		if (specialPaths.includes(entry.filename as typeof specialPaths[number])) {
			console.log(`${entry.filename} already written, skipping...`)
			continue
		}

		if (excludeSlidePaths.includes(entry.filename)) {
			console.log(`Excluding ${entry.filename}...`)
			continue
		}

		zipWriter.add(entry.filename, (await entry.getData(new BlobWriter())).stream())
	}

	return new Uint8Array(await (await zipWriter.close()).arrayBuffer())
}
