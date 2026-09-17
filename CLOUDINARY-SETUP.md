# Cloudinary reference uploads

This integration changes only Tudou and GRSAI reference uploads. Apilio, RH,
APIMart, model parameters, project originals, and existing ImgBB links are retained.
Tudou and GRSAI always use Cloudinary for new hosted uploads, even when old settings
contain an ImgBB preference. There is no ImgBB selector, key input, or fallback for
these two providers. Existing public ImgBB image links remain readable.

## Required server configuration

Create a dedicated **signed** Cloudinary upload preset for reference images.
Restrict allowed formats to `jpg,png,webp` and keep incoming/eager transformations
disabled. Verify account-side size and pixel limits through the usage API's
`media_limits`; the browser's 10,000,000-byte threshold is not a security boundary.
The preset API did not retain `max_file_size` during local setup, so do not assume
that sending this parameter enforces a preset-level size limit. The tested account
reported 10,485,760 bytes and 25,000,000 pixels on 2026-09-17.
Do not enable public unsigned uploads for this preset.
Leave folder/public-ID prefix rules unset so the fixed signed public ID is preserved.

Configure these environment variables on the server, never in the bundle:

| Variable | Value |
| --- | --- |
| `CLOUDINARY_CLOUD_NAME` | Product environment cloud name |
| `CLOUDINARY_API_KEY` | API key |
| `CLOUDINARY_API_SECRET` | Secret used only for signing |
| `CLOUDINARY_UPLOAD_PRESET` | Dedicated signed upload preset name |
| `CLOUDINARY_UPLOAD_TOKEN_HASHES` | Comma-separated SHA-256 hashes of authorized upload credentials |

Generate a random upload credential (at least 32 characters), hash it with SHA-256,
and configure only its hash on the server. Distribute the credential privately to
each authorized user, who enters it in the Cloudinary upload credential field.
The field is labelled `上传服务访问码`. This is an application-issued access code,
not an API key, API secret, or upload preset supplied by the Cloudinary dashboard.
Revoke a user by removing their hash. These credentials are distinct from provider
API keys and the Cloudinary API Secret.

The signing endpoint fails with 503 until configuration exists and with 401 for
unauthorized credentials. Signed parameters are fixed server-side, with unique
asset identifiers and `overwrite=false`. Images upload directly to Cloudinary:
neither image bodies nor Cloudinary secrets go through the front-end configuration.

The endpoint includes a 20-signature/minute/token **per-instance** burst limiter.
It is not a distributed quota, billing cap, or single-use-signature guarantee.
Before a public/shared-account rollout, configure platform-level distributed
rate limits and Cloudinary usage alerts. Cloudinary signatures have their own
validity period; removing a token stops new signatures, not already issued ones.

## Local and desktop routing

Web: `/api/cloudinary-signature` runs on the deployed server.
GitHub Pages requests the canonical `https://www.vinsen.top` signing endpoint.
The local Node server uses server environment credentials when configured.
Otherwise it forwards credentialed signing requests to the canonical hosted
endpoint. The local route does not accept arbitrary upstream URLs or signing
parameters. Server environment credentials are for local development only:
do not package a `.env` or secrets into a desktop installer.

Start a local preview with `PORT=4174 node local-preview-server.mjs`.
This source change does not update any existing Windows installer.

## Image behavior

Static JPEG, PNG, and WebP are supported. Compliant originals are uploaded byte
for byte. Over-limit copies target 9,500,000 bytes and at most 25,000,000 pixels:
quality is reduced at the original dimensions first, unless pixel limits already
require proportional resizing. WebP/PNG preserve alpha, never JPEG flattening.
Processing failures stop before model submission. Unsupported formats fail
explicitly; animations are not silently flattened.

Only request copies are encoded. No stored blob, canvas metadata, source file,
or saved project image is replaced by the upload result. Old trusted ImgBB and
Cloudinary links are reused. There is no persistent upload cache or cloud asset
deletion policy in this change; local original data remains authoritative.

Reference images are third-party-hosted public delivery assets. Use nonsensitive
images for initial testing. Account configuration and real-network model access
must be verified before deployment; mock tests do not prove provider acceptance.

## Reference

- https://cloudinary.com/documentation/upload_images
- https://cloudinary.com/documentation/upload_presets
- https://cloudinary.com/documentation/authentication_signatures
