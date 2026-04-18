import { CachedMetadata, moment, requestUrl, TFile, View, WorkspaceLeaf } from 'obsidian'
import { encryptString, sha1 } from './crypto'
import SharePlugin from './main'
import StatusMessage, { StatusType } from './StatusMessage'
import NoteTemplate, { ElementStyle, getElementStyle } from './NoteTemplate'
import { ThemeMode, TitleSource, YamlField } from './settings'
import { dataUriToBuffer } from 'data-uri-to-buffer'
import FileTypes from './libraries/FileTypes'
import { CheckFilesResult, parseExistingShareUrl } from './api'
import { minify } from 'csso'
import { InternalLinkMethod } from './types'
import DurationConstructor = moment.unitOfTime.DurationConstructor

const cssAttachmentWhitelist: { [key: string]: string[] } = {
  ttf: ['font/ttf', 'application/x-font-ttf', 'application/x-font-truetype', 'font/truetype'],
  otf: ['font/otf', 'application/x-font-opentype'],
  woff: ['font/woff', 'application/font-woff', 'application/x-font-woff'],
  woff2: ['font/woff2', 'application/font-woff2', 'application/x-font-woff2'],
  svg: ['image/svg+xml']
}

const shareNoteTocCss = `
.share-note-layout{display:block;}
.share-note-sidebar{margin:0 0 1.5rem;}
.share-note-content{min-width:0;}
.share-note-toc{position:relative;width:100%;padding:1rem;border:1px solid var(--background-modifier-border);border-radius:16px;background:var(--background-secondary);box-shadow:0 12px 30px rgba(15,23,42,.08);backdrop-filter:blur(10px);}
.share-note-toc-title{margin:0 0 .85rem;font-size:.72rem;font-weight:600;letter-spacing:.16em;text-transform:uppercase;color:var(--text-muted);}
.share-note-toc-list,.share-note-toc-children{margin:0;padding:0;list-style:none;}
.share-note-toc-list{display:flex;flex-direction:column;gap:.25rem;}
.share-note-toc-item{display:block;}
.share-note-toc-link{display:flex;align-items:center;gap:.7rem;padding:.45rem .55rem;border-radius:.8rem;color:inherit;text-decoration:none;transition:background-color 180ms ease,color 180ms ease,transform 180ms ease;}
.share-note-toc-link:hover{color:var(--text-accent);background:var(--background-primary-alt);}
.share-note-toc-link.active{color:var(--text-accent);background:var(--background-primary-alt);}
.share-note-toc-link.active .share-note-toc-badge{background:var(--text-accent);color:var(--background-primary);}
.share-note-toc-badge{display:inline-flex;align-items:center;justify-content:center;min-width:1.55rem;height:1.55rem;padding:0 .45rem;border-radius:999px;background:var(--interactive-accent);color:var(--text-on-accent, #fff);font-size:.72rem;font-weight:700;line-height:1;}
.share-note-toc-text{line-height:1.45;}
.share-note-toc-text-sub{font-size:.95em;font-weight:600;}
.share-note-toc-text-leaf{font-size:.84em;color:var(--text-muted);}
.share-note-toc-children{margin-left:1rem;padding-left:.9rem;border-left:1px dashed var(--background-modifier-border);max-height:0;opacity:0;overflow:hidden;will-change:max-height,opacity;transition:max-height 320ms cubic-bezier(.25,.8,.25,1),opacity 220ms ease-in-out;}
.share-note-toc-item:hover>.share-note-toc-children,.share-note-toc-item:focus-within>.share-note-toc-children,.share-note-toc-item.open>.share-note-toc-children{max-height:120vh;opacity:1;}
.share-note-toc-bootstrap{position:absolute;width:0;height:0;opacity:0;pointer-events:none;}
.share-note-toc-target{scroll-margin-top:1.5rem;}
@media (min-width: 1080px){
  .share-note-layout{display:grid;grid-template-columns:minmax(15rem,18rem) minmax(0,1fr);gap:clamp(1.5rem,3vw,2.75rem);align-items:start;}
  .share-note-sidebar{position:sticky;top:1.5rem;align-self:start;margin:0;}
}
`

const shareNoteTocBootstrap = `(function(trigger){var toc=trigger.closest('.share-note-toc');if(!toc||toc.dataset.shareNoteBound==='true')return;toc.dataset.shareNoteBound='true';var links=Array.from(toc.querySelectorAll('[data-toc-slug]'));var itemMap=new Map(Array.from(toc.querySelectorAll('[data-toc-item]')).map(function(item){return [item.getAttribute('data-toc-item')||'',item]}));var targets=links.map(function(link){var slug=link.getAttribute('data-toc-slug')||'';return {slug:slug,link:link,heading:document.getElementById(slug)};}).filter(function(entry){return entry.slug&&entry.heading;});var setActive=function(slug){links.forEach(function(link){var active=link.getAttribute('data-toc-slug')===slug;link.classList.toggle('active',active);if(active){link.setAttribute('aria-current','true');}else{link.removeAttribute('aria-current');}});itemMap.forEach(function(item){item.classList.remove('open');});var current=itemMap.get(slug||'');while(current){current.classList.add('open');current=current.parentElement&&current.parentElement.closest('[data-toc-item]');}};var computeActive=function(){if(!targets.length)return '';var offset=Math.min(window.innerHeight*.22,160);var current=targets[0].slug;for(var i=0;i<targets.length;i++){if(targets[i].heading.getBoundingClientRect().top-offset<=0){current=targets[i].slug;}else{break;}}return current;};var ticking=false;var update=function(){ticking=false;setActive(computeActive());};var schedule=function(){if(ticking)return;ticking=true;window.requestAnimationFrame(update);};window.addEventListener('scroll',schedule,{passive:true});window.addEventListener('resize',schedule,{passive:true});setTimeout(schedule,0);schedule();})(this)`

export interface SharedUrl {
  filename: string
  decryptionKey: string
  url: string
}

export interface SharedNote extends SharedUrl {
  file: TFile
}

export interface PreviewSection {
  el: HTMLElement
}

export interface Renderer {
  parsing: boolean,
  pusherEl: HTMLElement,
  previewEl: HTMLElement,
  sections: PreviewSection[]
}

interface TocHeading {
  level: number,
  text: string,
  target: string
}

interface TocNode extends TocHeading {
  children: TocNode[]
}

export interface ViewModes extends View {
  getViewType: any,
  getDisplayText: any,
  modes: {
    preview: {
      renderer: Renderer
    }
  }
}

export default class Note {
  plugin: SharePlugin
  leaf: WorkspaceLeaf
  status: StatusMessage
  css: string
  cssRules: CSSRule[]
  cssResult: CheckFilesResult['css']
  contentDom: Document
  meta: CachedMetadata | null
  isEncrypted = true
  isForceUpload = false
  isForceClipboard = false
  template: NoteTemplate
  elements: ElementStyle[]
  expiration?: number

  constructor (plugin: SharePlugin) {
    this.plugin = plugin
    // .getLeaf() doesn't return a `previewMode` property when a note is pinned,
    // so use the undocumented .getActiveFileView() which seems to work fine
    // @ts-ignore
    this.leaf = this.plugin.app.workspace.getActiveFileView()?.leaf
    this.elements = []
    this.template = new NoteTemplate()
  }

  /**
   * Return the name (key) of a frontmatter property, eg 'share_link'
   * @param key
   * @return {string} The name (key) of a frontmatter property
   */
  field (key: YamlField): string {
    return this.plugin.field(key)
  }

  async share () {
    if (!this.plugin.settings.apiKey) {
      this.plugin.authRedirect('share').then()
      return
    }

    // Create a semi-permanent status notice which we can update
    this.status = new StatusMessage('Please do not change to another note as the current note data is still being parsed.', StatusType.Default, 60 * 1000)

    // Switch to reading mode
    const startMode = this.leaf.getViewState()
    const previewMode = this.leaf.getViewState()
    if (previewMode.state) {
      previewMode.state.mode = 'preview'
    }
    await this.leaf.setViewState(previewMode)
    // Add a delay to wait for reading mode to finalise rendering - https://github.com/alangrainger/share-note/discussions/162#discussioncomment-15394971
    await new Promise(resolve => setTimeout(resolve, 600))

    // Scroll the view to the top to ensure we get the default margins for .markdown-preview-pusher
    // @ts-ignore
    this.leaf.view.previewMode.applyScroll(0) // 'view.previewMode'
    await new Promise(resolve => setTimeout(resolve, 100))
    try {
      const view = this.leaf.view as ViewModes
      const renderer = view.modes.preview.renderer
      // Copy classes and styles
      this.elements.push(getElementStyle('html', document.documentElement))
      const bodyStyle = getElementStyle('body', document.body)
      bodyStyle.classes.push('share-note-plugin') // Add a targetable class for published notes
      this.elements.push(bodyStyle)
      this.elements.push(getElementStyle('preview', renderer.previewEl))
      this.elements.push(getElementStyle('pusher', renderer.pusherEl))
      this.contentDom = new DOMParser().parseFromString(await this.querySelectorAll(this.leaf.view as ViewModes), 'text/html')
      this.cssRules = []
      Array.from(document.styleSheets)
        .forEach(x => Array.from(x.cssRules)
          .forEach(rule => {
            this.cssRules.push(rule)
          }))

      // Merge all CSS rules into a string for later minifying
      this.css = this.cssRules
        .filter((rule: CSSMediaRule) => {
          /*
          Remove styles that prevent a print preview from showing on the web, thanks to @texastoland on Github
          https://github.com/alangrainger/share-note/issues/75#issuecomment-2708719828

          This removes all "@media print" rules, which in my testing doesn't appear to have any negative effect.
          Will have to revisit this if users discover issues.
          */
          return rule?.media?.[0] !== 'print'
        })
        .map(rule => rule.cssText).join('').replace(/\n/g, '')
    } catch (e) {
      console.log(e)
      this.status.hide()
      new StatusMessage('Failed to parse current note, check console for details', StatusType.Error)
      return
    }

    // Reset the view to the original mode
    // The timeout is required, even though we 'await' the preview mode setting earlier
    setTimeout(() => {
      this.leaf.setViewState(startMode)
    }, 200)

    this.status.setStatus('Processing note...')
    const file = this.plugin.app.workspace.getActiveFile()
    if (!(file instanceof TFile)) {
      // No active file
      this.status.hide()
      new StatusMessage('There is no active file to share')
      return
    }
    this.meta = this.plugin.app.metadataCache.getFileCache(file)

    // Generate the HTML file for uploading

    if (this.plugin.settings.removeYaml) {
      // Remove frontmatter to avoid sharing unwanted data
      this.contentDom.querySelector('div.metadata-container')?.remove()
      this.contentDom.querySelector('pre.frontmatter')?.remove()
      this.contentDom.querySelector('div.frontmatter-container')?.remove()
    } else {
      // Frontmatter properties are weird - the DOM elements don't appear to contain any data.
      // We get the property name from the data-property-key and set that on the labelEl value,
      // then take the corresponding value from the metadataCache and set that on the valueEl value.
      this.contentDom.querySelectorAll('div.metadata-property')
        .forEach(propertyContainerEl => {
          const propertyName = propertyContainerEl.getAttribute('data-property-key')
          if (propertyName) {
            const labelEl = propertyContainerEl.querySelector('input.metadata-property-key-input')
            labelEl?.setAttribute('value', propertyName)
            const valueEl = propertyContainerEl.querySelector('div.metadata-property-value > input')
            const value = this.meta?.frontmatter?.[propertyName] || ''
            valueEl?.setAttribute('value', value)
            // Special cases for different element types
            switch (valueEl?.getAttribute('type')) {
              case 'checkbox':
                if (value) valueEl.setAttribute('checked', 'checked')
                break
            }
          }
        })
    }
    if (this.plugin.settings.removeBacklinksFooter) {
      // Remove backlinks footer
      this.contentDom.querySelector('div.embedded-backlinks')?.remove()
    } else {
      // Make backlinks clickable
      for (const el of this.contentDom.querySelectorAll<HTMLElement>('.embedded-backlinks .search-result-file-title.is-clickable')) {
        // Get the inner text, which is the name of the destination note
        const linkText = (el.querySelector('.tree-item-inner') as HTMLElement)?.innerText
        // Replace with a clickable link if possible
        if (linkText) this.internalLinkToSharedNote(linkText, el, InternalLinkMethod.ONCLICK)
      }
    }

    // Fix callout icons
    const defaultCalloutType = this.getCalloutIcon(selectorText => selectorText === '.callout') || 'pencil'
    for (const el of this.contentDom.getElementsByClassName('callout')) {
      // Get the callout icon from the CSS. I couldn't find any way to do this from the DOM,
      // as the elements may be far down below the fold and are not populated.
      const type = el.getAttribute('data-callout')
      let icon = this.getCalloutIcon(selectorText => selectorText.includes(`data-callout="${type}"`)) || defaultCalloutType
      icon = icon.replace('lucide-', '')
      // Replace the existing icon so we:
      // a) don't get double-ups, and
      // b) have a consistent style
      const iconEl = el.querySelector('div.callout-icon')
      const svgEl = iconEl?.querySelector('svg')
      if (svgEl) {
        svgEl.outerHTML = `<svg width="16" height="16" data-share-note-lucide="${icon}"></svg>`
      }
    }

    // Replace links
    for (const el of this.contentDom.querySelectorAll<HTMLElement>('a.internal-link, a.footnote-link')) {
      const href = el.getAttribute('href')
      const match = href ? href.match(/^([^#]+)/) : null
      if (href?.match(/^#/)) {
        // This is an Anchor link to a document heading, we need to add custom Javascript
        // to jump to that heading rather than using the normal # link
        try {
          this.setHeadingAnchorAction(el, href.slice(1))
          continue
        } catch (e) {
          console.error(e)
        }
      } else if (match) {
        if (this.internalLinkToSharedNote(match[1], el)) {
          // The internal link could be linked to another shared note
          continue
        }
      }
      // This linked note is not shared, so remove the link and replace with the non-link content
      el.replaceWith(el.innerText)
    }

    // Remove target=_blank from external links
    this.contentDom
      .querySelectorAll<HTMLElement>('a.external-link')
      .forEach(el => el.removeAttribute('target'))

    // Remove elements by user's custom CSS selectors (if any)
    this.plugin.settings.removeElements
      .split('\n').map(s => s.trim()).filter(Boolean)
      .forEach(selector => this.contentDom.querySelectorAll(selector)
        .forEach(el => el.remove()))

    this.injectTableOfContents()

    // Note options
    this.expiration = this.getExpiration()

    // Process CSS and images
    const uploadResult = await this.processMedia()
    this.cssResult = uploadResult.css
    await this.processCss()

    /*
     * Encrypt the note contents
     */

    // Use previous name and key if they exist, so that links will stay consistent across updates
    let decryptionKey = ''
    if (this.meta?.frontmatter?.[this.field(YamlField.link)]) {
      const match = parseExistingShareUrl(this.meta?.frontmatter?.[this.field(YamlField.link)])
      if (match) {
        this.template.filename = match.filename
        decryptionKey = match.decryptionKey
      }
    }
    this.template.encrypted = this.isEncrypted

    // Select which source for the title
    let title
    switch (this.plugin.settings.titleSource) {
      case TitleSource['First H1']:
        title = this.contentDom.getElementsByTagName('h1')?.[0]?.innerText
        break
      case TitleSource['Frontmatter property']:
        title = this.meta?.frontmatter?.[this.field(YamlField.title)]
        break
    }
    if (!title) {
      // Fallback to basename if either of the above fail
      title = file.basename
    }

    if (this.isEncrypted) {
      this.status.setStatus('Encrypting note...')
      const plaintext = JSON.stringify({
        content: this.contentDom.body.innerHTML,
        basename: title
      })
      // Encrypt the note
      const encryptedData = await encryptString(plaintext, decryptionKey)
      this.template.content = JSON.stringify({
        ciphertext: encryptedData.ciphertext
      })
      decryptionKey = encryptedData.key
    } else {
      // This is for notes shared without encryption, using the
      // share_unencrypted frontmatter property
      this.template.content = this.contentDom.body.innerHTML
      this.template.title = title
      // Create a meta description preview based off the <p> elements
      const desc = Array.from(this.contentDom.querySelectorAll('p'))
        .map(x => x.innerText).filter(x => !!x)
        .join(' ')
      this.template.description = desc.length > 200 ? desc.slice(0, 197) + '...' : desc
    }

    // Make template value replacements
    this.template.width = this.plugin.settings.noteWidth
    // Set theme light/dark
    if (this.plugin.settings.themeMode !== ThemeMode['Same as theme']) {
      this.elements
        .filter(x => x.element === 'body')
        .forEach(item => {
          // Remove the existing theme setting
          item.classes = item.classes.filter(cls => cls !== 'theme-dark' && cls !== 'theme-light')
          // Add the preferred theme setting (dark/light)
          item.classes.push('theme-' + ThemeMode[this.plugin.settings.themeMode].toLowerCase())
        })
    }
    this.template.elements = this.elements
    // Check for MathJax
    this.template.mathJax = !!this.contentDom.body.innerHTML.match(/<mjx-container/)

    // Share the file
    this.status.setStatus('Uploading note...')
    let shareLink = await this.plugin.api.createNote(this.template, this.expiration)
    requestUrl(shareLink).then().catch() // Fetch the uploaded file to pull it through the cache

    // Add the decryption key to the share link
    if (shareLink && this.isEncrypted) {
      shareLink += '#' + decryptionKey
    }

    let shareMessage = 'The note has been shared'
    if (shareLink) {
      await this.plugin.app.fileManager.processFrontMatter(file, (frontmatter) => {
        // Update the frontmatter with the share link
        frontmatter[this.field(YamlField.link)] = shareLink
        frontmatter[this.field(YamlField.updated)] = moment().format()
      })
      if (this.plugin.settings.clipboard || this.isForceClipboard) {
        // Copy the share link to the clipboard
        try {
          await navigator.clipboard.writeText(shareLink)
          shareMessage = `${shareMessage} and the link is copied to your clipboard 📋`
        } catch (e) {
          // If there's an error here it's because the user clicked away from the Obsidian window
        }
        this.isForceClipboard = false
      }
    }

    this.status.hide()
    new StatusMessage(shareMessage + `<br><br><a href="${shareLink}">↗️ Open shared note</a>`, StatusType.Success, 6000)
  }

  /**
   * Upload media attachments
   */
  async processMedia () {
    const elements = ['img', 'video']
    this.status.setStatus('Processing attachments...')
    for (const el of this.contentDom.querySelectorAll(elements.join(','))) {
      const src = el.getAttribute('src')
      if (!src) continue
      let content, filetype

      if (src.startsWith('http') && !src.match(/^https?:\/\/localhost/)) {
        // This is a web asset, no need to upload
        continue
      }

      const filesource = el.getAttribute('filesource')
      if (filesource?.match(/excalidraw/i)) {
        // Excalidraw drawing
        console.log('Processing Excalidraw drawing...')
        try {
          // @ts-ignore
          const excalidraw = this.plugin.app.plugins.getPlugin('obsidian-excalidraw-plugin')
          if (!excalidraw) continue
          content = await excalidraw.ea.createSVG(filesource)
          content = content.outerHTML
          filetype = 'svg'
          /*
            Or as PNG:
            const blob = await excalidraw.ea.createPNG(filesource)
            content = await blob.arrayBuffer()
            filetype = 'png'
          */
        } catch (e) {
          console.error('Unable to process Excalidraw drawing:')
          console.error(e)
        }
      } else {
        try {
          const res = await fetch(src)
          if (res && res.status === 200) {
            content = await res.arrayBuffer()
            const parsed = new URL(src)
            filetype = parsed.pathname.split('.').pop()
          }
        } catch (e) {
          // Unable to process this file
          continue
        }
      }

      if (filetype && content) {
        const hash = await sha1(content)
        await this.plugin.api.queueUpload({
          data: {
            filetype,
            hash,
            content,
            byteLength: content.byteLength,
            expiration: this.expiration
          },
          callback: (url) => el.setAttribute('src', url)
        })
      }
      el.removeAttribute('alt')
    }
    return this.plugin.api.processQueue(this.status)
  }

  /**
   * Upload theme CSS, unless this file has previously been shared,
   * or the user has requested a force re-upload
   */
  async processCss () {
    // Upload the main CSS file only if the user has asked for it.
    // We do it this way to ensure that the CSS the user wants on the server
    // stays that way, until they ASK to overwrite it.
    if (this.isForceUpload || !this.cssResult) {
      // Extract any attachments from the CSS.
      // Will use the mime-type whitelist to determine which attachments to extract.
      this.status.setStatus('Processing CSS...')
      const attachments = this.css.match(/url\s*\(.*?\)/g) || []
      for (const attachment of attachments) {
        const assetMatch = attachment.match(/url\s*\(\s*"*(.*?)\s*(?<!\\)"\s*\)/)
        if (!assetMatch) continue
        const assetUrl = assetMatch?.[1] || ''
        if (assetUrl.startsWith('data:')) {
          // Attempt to parse the data URL
          const parsed = dataUriToBuffer(assetUrl)
          if (parsed?.type) {
            if (parsed.type === 'application/octet-stream') {
              // Attempt to get type from magic bytes
              const decoded = FileTypes.getFromSignature(parsed.buffer)
              if (!decoded) continue
              parsed.type = decoded.mimetype
            }
            const filetype = this.extensionFromMime(parsed.type)
            if (filetype) {
              const hash = await sha1(parsed.buffer)
              await this.plugin.api.queueUpload({
                data: {
                  filetype,
                  hash,
                  content: parsed.buffer,
                  byteLength: parsed.buffer.byteLength,
                  expiration: this.expiration
                },
                callback: (url) => {
                  this.css = this.css.replace(assetMatch[0], `url("${url}")`)
                }
              })
            }
          }
        } else if (assetUrl && !assetUrl.startsWith('http')) {
          // Locally stored CSS attachment
          const filename = assetUrl.match(/([^/\\]+)\.(\w+)$/)
          if (filename) {
            if (cssAttachmentWhitelist[filename[2]]) {
              // Fetch the attachment content
              const res = await fetch(assetUrl)
              // Reupload to the server
              const contents = await res.arrayBuffer()
              const hash = await sha1(contents)
              await this.plugin.api.queueUpload({
                data: {
                  filetype: filename[2],
                  hash,
                  content: contents,
                  byteLength: contents.byteLength,
                  expiration: this.expiration
                },
                callback: (url) => {
                  this.css = this.css.replace(assetMatch[0], `url("${url}")`)
                }
              })
            }
          }
        }
      }
      this.status.setStatus('Uploading CSS attachments...')
      await this.plugin.api.processQueue(this.status, 'CSS attachment')
      this.status.setStatus('Uploading CSS...')
      const minified = minify(this.css).css
      const cssHash = await sha1(minified)
      try {
        if (cssHash !== this.cssResult?.hash) {
          await this.plugin.api.upload({
            filetype: 'css',
            hash: cssHash,
            content: minified,
            byteLength: minified.length,
            expiration: this.expiration
          })
        }

        // Store the CSS theme in the settings
        // @ts-ignore
        this.plugin.settings.theme = this.plugin.app?.customCss?.theme || '' // customCss is not exposed
        await this.plugin.saveSettings()
      } catch (e) {
      }
    }
  }

  async querySelectorAll (view: ViewModes) {
    const renderer = view.modes.preview.renderer
    let html = ''
    await new Promise<void>(resolve => {
      let count = 0
      let parsing = 0
      const timer = setInterval(() => {
        try {
          const sections = renderer.sections
          count++
          if (renderer.parsing) parsing++
          if (count > parsing) {
            // Check the final sections to see if they have rendered
            let rendered = 0
            if (sections.length > 12) {
              sections.slice(sections.length - 7, sections.length - 1).forEach((section: PreviewSection) => {
                if (section.el.innerHTML) rendered++
              })
              if (rendered > 3) count = 100
            } else {
              count = 100
            }
          }
          if (count > 40) {
            html = this.reduceSections(renderer.sections)
            resolve()
          }
        } catch (e) {
          clearInterval(timer)
          resolve()
        }
      }, 100)
    })
    return html
  }

  /**
   * Takes a linkText like 'Some note' or 'Some path/Some note.md' and sees if that note is already shared.
   * If it's already shared, then replace the internal link with the public link to that note.
   */
  internalLinkToSharedNote (linkText: string, el: HTMLElement, method: InternalLinkMethod = 0) {
    try {
      // This is an internal link to another note - check to see if we can link to an already shared note
      const linkedFile = this.plugin.app.metadataCache.getFirstLinkpathDest(linkText, '')
      if (linkedFile instanceof TFile) {
        const linkedMeta = this.plugin.app.metadataCache.getFileCache(linkedFile)
        const href = linkedMeta?.frontmatter?.[this.field(YamlField.link)]
        if (href && typeof href === 'string') {
          // This file is shared, so update the link with the share URL
          if (method === InternalLinkMethod.ANCHOR) {
            // Set the href for an <a> element
            el.setAttribute('href', href)
            el.removeAttribute('target')
          } else if (method === InternalLinkMethod.ONCLICK) {
            // Add an onclick() method
            el.setAttribute('onclick', `window.location.href='${href}'`)
            el.classList.add('force-cursor')
          }
          return true
        }
      }
    } catch (e) {
      console.error(e)
    }
    return false
  }

  getCalloutIcon (test: (selectorText: string) => boolean) {
    const rule = this.cssRules
      .find((rule: CSSStyleRule) => rule.selectorText && test(rule.selectorText) && rule.style.getPropertyValue('--callout-icon')) as CSSStyleRule
    if (rule) {
      return rule.style.getPropertyValue('--callout-icon')
    }
    return ''
  }

  reduceSections (sections: { el: HTMLElement }[]) {
    return sections.reduce((p: string, c) => p + c.el.outerHTML, '')
  }

  injectTableOfContents () {
    const headings = this.collectTocHeadings()
    const tocTree = this.buildTocTree(headings)
    if (!tocTree.length) return

    const layout = this.contentDom.createElement('div')
    layout.classList.add('share-note-layout')

    const sidebar = this.contentDom.createElement('aside')
    sidebar.classList.add('share-note-sidebar')

    const nav = this.contentDom.createElement('nav')
    nav.classList.add('share-note-toc')
    nav.setAttribute('aria-label', 'Table of contents')

    const title = this.contentDom.createElement('div')
    title.classList.add('share-note-toc-title')
    title.innerText = 'On this page'
    nav.append(title)
    nav.append(this.renderTocTree(tocTree))

    const bootstrap = this.contentDom.createElement('img')
    bootstrap.classList.add('share-note-toc-bootstrap')
    bootstrap.setAttribute('src', 'data:,')
    bootstrap.setAttribute('alt', '')
    bootstrap.setAttribute('aria-hidden', 'true')
    bootstrap.setAttribute('onload', shareNoteTocBootstrap)
    nav.append(bootstrap)

    sidebar.append(nav)

    const content = this.contentDom.createElement('div')
    content.classList.add('share-note-content')
    Array.from(this.contentDom.body.childNodes).forEach(node => content.append(node))

    layout.append(sidebar, content)
    this.contentDom.body.append(layout)
    this.css += shareNoteTocCss
  }

  collectTocHeadings () {
    const usedTargets = new Set<string>()

    return Array.from(this.contentDom.body.querySelectorAll<HTMLElement>('h1, h2, h3, h4, h5, h6'))
      .filter(el => !el.closest('.share-note-toc'))
      .map(el => {
        const text = el.innerText.trim()
        if (!text) return null

        const level = parseInt(el.tagName.slice(1), 10)
        if (level < 2 || level > 4) return null
        const existingTarget = el.getAttribute('id') || el.getAttribute('data-heading')
        const target = existingTarget || this.getUniqueHeadingTarget(text, usedTargets)

        if (!el.getAttribute('data-heading')) {
          el.setAttribute('data-heading', target)
        }
        if (!el.getAttribute('id')) {
          el.setAttribute('id', target)
        }
        el.classList.add('share-note-toc-target')
        usedTargets.add(target)

        return {
          level,
          text,
          target
        } as TocHeading
      })
      .filter((heading): heading is TocHeading => !!heading)
  }

  buildTocTree (headings: TocHeading[]) {
    const tocTree: TocNode[] = []
    const parents: TocNode[] = []

    headings.forEach(heading => {
      const node: TocNode = {
        ...heading,
        children: []
      }

      if (node.level === 2) {
        tocTree.push(node)
        parents[0] = node
        parents.length = 1
        return
      }

      if (node.level === 3 && parents[0]) {
        parents[0].children.push(node)
        parents[1] = node
        parents.length = 2
        return
      }

      if (node.level === 4 && parents[1]) {
        parents[1].children.push(node)
        parents[2] = node
        parents.length = 3
      }
    })

    return tocTree
  }

  renderTocTree (nodes: TocNode[], parents: number[] = []) {
    const list = this.contentDom.createElement('ul')
    list.classList.add(parents.length ? 'share-note-toc-children' : 'share-note-toc-list')

    nodes.forEach((node, index) => {
      const order = [...parents, index + 1]
      const item = this.contentDom.createElement('li')
      item.classList.add('share-note-toc-item')
      item.setAttribute('data-toc-item', node.target)

      const link = this.contentDom.createElement('a')
      link.classList.add('share-note-toc-link', 'internal-link')
      link.setAttribute('href', '#' + node.target)
      link.setAttribute('data-toc-slug', node.target)

      if (node.level === 2) {
        const badge = this.contentDom.createElement('span')
        badge.classList.add('share-note-toc-badge')
        badge.innerText = order[0].toString()
        link.append(badge)
      }

      const text = this.contentDom.createElement('span')
      text.classList.add('share-note-toc-text')
      if (node.level === 3) {
        text.classList.add('share-note-toc-text-sub')
        text.innerText = `${order.join('.')} ${node.text}`
      } else if (node.level === 4) {
        text.classList.add('share-note-toc-text-leaf')
        text.innerText = `${order.join('.')} ${node.text}`
      } else {
        text.innerText = node.text
      }
      link.append(text)

      this.setHeadingAnchorAction(link, node.target)
      item.append(link)

      if (node.children.length) {
        item.append(this.renderTocTree(node.children, order))
      }

      list.append(item)
    })

    return list
  }

  getUniqueHeadingTarget (text: string, usedTargets: Set<string>) {
    let target = text
    let suffix = 2

    while (usedTargets.has(target)) {
      target = `${text}-${suffix++}`
    }

    return target
  }

  setHeadingAnchorAction (el: HTMLElement, heading: string) {
    const escapedHeading = heading.replace(/(['"])/g, '\\$1')
    const linkTypes = [
      `[data-heading="${escapedHeading}"]`,
      `[id="${escapedHeading}"]`,
    ]

    linkTypes.forEach(selector => {
      if (this.contentDom.querySelectorAll(selector)?.[0]) {
        // Double-escape the double quotes (but leave single quotes single escaped)
        // It makes sense if you look at the query selector...
        el.setAttribute('onclick', `document.querySelectorAll('${selector.replace(/"/g, '\\"')}')[0].scrollIntoView(true)`)
      }
    })
    el.removeAttribute('target')
    el.removeAttribute('href')
  }

  /**
   * Turn the font mime-type into an extension.
   * @param {string} mimeType
   * @return {string|undefined}
   */
  extensionFromMime (mimeType: string): string | undefined {
    const mimes = cssAttachmentWhitelist
    return Object.keys(mimes).find(x => mimes[x].includes((mimeType || '').toLowerCase()))
  }

  /**
   * Get the value of a frontmatter property
   */
  getProperty (field: YamlField) {
    return this.meta?.frontmatter?.[this.plugin.field(field)]
  }

  /**
   * Force all related assets to upload again
   */
  forceUpload () {
    this.isForceUpload = true
  }

  /**
   * Copy the shared link to the clipboard, regardless of the user setting
   */
  forceClipboard () {
    this.isForceClipboard = true
  }

  /**
   * Enable/disable encryption for the note
   */
  shareAsPlainText (isPlainText: boolean) {
    this.isEncrypted = !isPlainText
  }

  /**
   * Calculate an expiry datetime from the provided expiry duration
   */
  getExpiration () {
    const whitelist = ['minute', 'hour', 'day', 'month']
    const expiration = this.getProperty(YamlField.expires) || this.plugin.settings.expiry
    if (expiration) {
      // Check for sanity against expected format
      const match = expiration.match(/^(\d+) ([a-z]+?)s?$/)
      if (match && whitelist.includes(match[2])) {
        return parseInt(moment().add(+match[1], (match[2] + 's') as DurationConstructor).format('x'), 10)
      }
    }
  }
}
