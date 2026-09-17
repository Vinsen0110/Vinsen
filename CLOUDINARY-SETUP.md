# Personal Cloudinary reference uploads

Only Tudou and GRSAI use this integration. Each user registers their own Cloudinary
account and enters their own Cloud Name, API Key and API Secret in API settings.
There is no shared account, application access code, Vercel environment variable,
server signing endpoint or image-host fallback to ImgBB.

## Account setup

1. Register a Cloudinary account and open its API credentials settings.
2. Copy that account's Cloud Name, API Key and API Secret into this app's
   Cloudinary fields. API Key and API Secret are masked in the UI.
3. Use an account/key authorized for signed image uploads. No unsigned preset or
   dedicated signed preset is required by this integration.
4. Check the account's default upload preset and remove any unwanted incoming
   transformations if original dimensions and transparency must be retained.

The account owns its uploaded images, storage, bandwidth limits and any charges.
This integration does not configure billing, quotas, automatic deletion or
Cloudinary usage alerts. Do not share one account's credentials with other users.

## Credential boundary

This is a bring-your-own-key browser client. Credentials are retained in the
user's existing local browser app settings, not in Vercel, the source bundle,
project image files or model-provider request bodies.

The browser creates a SHA-256 upload signature with Web Crypto. Only the API Key,
signature, timestamp, unique public ID, overwrite=false and image are sent
directly to Cloudinary. The API Secret is used locally and is not sent to a
signing server, Cloudinary upload endpoint or model provider.

Local browser storage is NOT an encrypted credential vault. Scripts running on
the same origin, browser extensions with access, or another person using that
browser profile may access saved credentials. Only enter your own credentials
on a trusted installation; do not use this model to distribute an operator-owned
secret. This personal-account mode is not equivalent security to a backend-held
secret.

Changing accounts applies to new uploads. An upload already in progress keeps
the account snapshot it started with. Existing public Cloudinary and ImgBB URLs
remain usable and are never moved to a different account automatically.

## Image behavior

Static JPEG, PNG and WebP are supported. Compliant originals are uploaded byte
for byte. Oversized request copies target 9,500,000 bytes and at most 25,000,000
pixels. Quality is reduced at original dimensions first, unless pixel limits
already require proportional resizing. WebP/PNG preserve alpha.

The browser starts compressing above 10,000,000 bytes. Actual limits depend on
each user's Cloudinary account; this check is not an account-side security or
billing limit. Unsupported formats and failed compression stop before model
submission. No original file, stored blob, canvas image or project metadata is
replaced by the uploaded copy. There is no persistent upload cache.

RH, Apilio and APIMart retain their native upload implementations and model
parameters. This source change does not build or update Windows installers.

## Local preview

Run `PORT=4186 node local-preview-server.mjs` without Cloudinary environment
variables. Browser signing requires HTTPS or a secure localhost context.

## References

- https://cloudinary.com/documentation/authentication_signatures
- https://cloudinary.com/documentation/upload_images
