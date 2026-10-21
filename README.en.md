# PPT Translation Tool

[简体中文](README.md) · [Live tool](https://ppt.naodong.store/) · [MIT license](LICENSE)

A browser-based assistant for translating PPTX text with an AI service of your choice. Extract paragraphs, copy numbered prompts, paste translations, and export a new presentation. PPTX parsing and writeback happen locally in your browser; the app has no built-in translation model and requires no OpenAI API key.

## Features

- Extract paragraphs with stable per-slide IDs and copy prompts in full or in batches.
- Preserve rich-text styles using validated markers that can move with translated phrases.
- Preserve manual breaks and dynamic fields; reject malformed, duplicate, or unknown markers before export.
- Export with the original formatting, or opt into browser-measured font fitting with explicit size limits and warnings.
- Light and dark themes; deploy as a static website.

## Run locally

Use Node.js 22.12+ (or Node.js 20.19+, as supported by Vite 7).

```sh
npm ci
npm run dev
```

Run `npm test` for regression checks and `npm run build` to typecheck and build into `dist/`. Serve that directory using a static host.

## Privacy and analytics

Presentation files, extracted content, and translations are processed locally. Text you manually copy into an external AI service, including the filename in the generated prompt, is subject to that service's policies.

Analytics is disabled by default for local and self-hosted builds. Optional settings are documented in [.env.example](.env.example). Only configure a service you operate. All `VITE_*` values are public browser configuration, never secrets. Rebuild after changing them.

When enabled, analytics sends the site ID, page URL including query parameters, referrer, user agent, and a visitor ID stored in localStorage. It does not send document text or translations. Hosting services may also log request IP addresses. The original hosted site uses analytics.

## Limitations

Only modern `.pptx` files are supported. Image-only text is not translated. Copying text to an external AI is a manual step. A valid marker structure does not guarantee translation quality or pixel-identical layout across PowerPoint, WPS, and fonts. Review exported files in your presentation software. Smart fitting does not enlarge text boxes or truncate text; unsupported or overflowing layouts produce warnings.

## Contributing and license

Report issues or send focused pull requests on GitHub. Include reproduction steps and your presentation software version. Use synthetic minimal examples; do not upload private or third-party documents without permission.

Project code is licensed under [MIT](LICENSE). Third-party dependencies retain their own licenses; see [third-party notices](public/THIRD_PARTY_NOTICES.txt).

## Public release maintenance

This repository is a complete, independently buildable MIT-licensed release. Day-to-day development takes place in a separate development repository. The maintainer periodically publishes reviewed stable updates here, preserving this public repository's release history. Issues and pull requests are welcome. Updating this repository does not automatically deploy the hosted website.
