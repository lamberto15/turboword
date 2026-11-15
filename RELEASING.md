# Releasing

`turboword` is released with [Changesets](https://github.com/changesets/changesets) and the
[`release.yml`](./.github/workflows/release.yml) workflow. After the first version, nothing is
published from a developer machine.

## Day to day

For every user-visible change (feature, fix, behavior change), add a changeset to your PR:

```sh
pnpm changeset
```

Pick the bump (`patch` for fixes, `minor` for features, `major` for breaking changes; while the
version is `0.x`, breaking changes are `minor`), write a one-line summary for the changelog, and
commit the generated `.changeset/*.md`. Docs-, test- and CI-only PRs don't need one (label them
`no-changeset`).

## How a release happens

1. PRs with changesets are merged into `main`.
2. `release.yml` opens or updates a **"chore(release): version packages"** PR that bumps
   `package.json` and writes `CHANGELOG.md`. Nothing is published yet.
3. Merging that PR runs the publish job: `pnpm run verify` (the same checks as CI), the package
   checks (`check:package`, `size`), then `changeset publish` publishes to npm through **trusted
   publishing** (OIDC: no npm token exists), with a provenance attestation, and creates the git
   tag and GitHub release.

Before merging the version PR, check that `CHANGELOG.md` reads well and that CI is green.

## One-time setup (maintainer)

### Claim the names on npm

Checked on 2026-11-14: the package name `turboword` was unpublished (`npm view turboword` → E404,
and the similar `turbo-word`, `turbo_word`, `turbo.word` were free too), and the scope
`@turboword` was free (`https://registry.npmjs.org/-/org/turboword/package` → 404 "Scope not
found"). Nothing reserves a name until it is taken, so do both soon:

- [ ] npm account with **2FA** enabled (security key or passkey recommended).
- [ ] **Reserve the `@turboword` organization** for the planned `@turboword/fonts` and
  `@turboword/harfbuzz` packages: npmjs.com → your avatar → **Add Organization** → name
  `turboword`, plan **Unlimited public packages** (free). npm users and organizations share one
  namespace, so creating the organization is what claims the scope. Afterwards,
  `npm org ls turboword` lists you as owner.
- [ ] **Claim `turboword`** with the first publish below.

### First publish, by hand, in your own terminal

npm only allows a trusted publisher on a package that already exists, so the first version is
published manually. **Run this yourself in a real terminal; an AI agent can't** (not even through
an agent's shell prefix): with 2FA by security key or passkey, `pnpm publish` fails without a TTY
(`ERR_PNPM_OTP_NON_INTERACTIVE`) and `npm publish --auth-type=web` prints the login URL and exits
with `EOTP` instead of waiting.

Decide first which version `latest` should point to. The first version of a new package always
gets `latest`, even with `--tag`. Releases of this package go to `latest` (Changesets is not in
pre mode), so the recommended route is: publish the current `0.0.0` from `main` now to claim the
name; the first CI release (for example `0.1.0`) then takes `latest` over automatically, and you
deprecate the placeholder with `npm deprecate turboword@0.0.0 "placeholder, use >=0.1.0"`. If you
would rather the first version on npm be a real one, merge the first version PR instead, let its
Release run fail at publish (expected, see below), and publish that version by hand with the same
commands.

```sh
git switch main && git pull
pnpm install --frozen-lockfile
pnpm run verify
pnpm run check:package
rm -rf /tmp/first-publish && mkdir /tmp/first-publish
pnpm pack --pack-destination /tmp/first-publish
tar -tzf /tmp/first-publish/*.tgz      # check the file list: dist/, README, both licences, NOTICE
npm login
npm publish /tmp/first-publish/*.tgz --access public --auth-type=web
```

Any Release run on `main` before this point fails at publish because the package doesn't exist
yet and no trusted publisher is configured. That is expected and publishes nothing.

After publishing, **wait, don't retry**: npm prints "may take a few minutes", the package can take
several minutes to appear on the registry, and an interim `0.0.0-stage` version may show up first.

- [ ] First publish done (commands above).
- [ ] `latest` points where it should: after the first CI release, `npm view turboword dist-tags`
  shows the CI version as `latest`, then `npm deprecate turboword@0.0.0 "placeholder, use >=0.1.0"`
  (skip if the first publish was a real version).

### Trusted publishing (configure right before the first CI release)

npm doesn't validate these values when you save them, and a new trusted publisher **expires if no
publish succeeds within 2 days** (npm docs, November 2026). Re-running a Release whose version is
already on npm publishes nothing and doesn't count. So configure it when the first "version
packages" PR is ready, then merge that PR.

- [ ] npmjs.com → `turboword` → **Settings → Trusted publishing** → GitHub Actions:
  owner `lamberto15`, repository `turboword`, workflow `release.yml` (file name only, exact,
  case-sensitive), environment: leave empty.
- [ ] Same form, **Allowed actions**: also allow direct **`npm publish`**. Publishers created
  after 2026-09-03 allow only `npm stage publish` by default, but `changeset publish` publishes
  directly (with `pnpm publish`), so a stage-only publisher can't release.
- [ ] Same page: publishing access → **require 2FA and disallow tokens**.

### GitHub

- [ ] Create the public repository `lamberto15/turboword` and push `main` (`main` is the default
  branch that `ci.yml`, `release.yml` and `.changeset/config.json` expect).
- [ ] GitHub → repo Settings → Actions → General: workflow permissions **Read**, and tick
  **"Allow GitHub Actions to create and approve pull requests"** (personal account: no
  organization-level setting to change first).
- [ ] Optional: repository secret `RELEASE_PAT` (fine-grained, this repo only, Contents and
  Pull requests: read and write). Without it the version PR shows "no checks reported" and is
  merged through the admin bypass; that is safe because CI already ran on `main` for every
  change in it and the publish job runs `pnpm run verify` before publishing. With it the version
  PR gets normal CI, but there is a token to store and renew.
- [ ] Branch ruleset on `main` requiring a PR and the **`ci-ok`** check (the only required check;
  it aggregates the Node 22/24/26 × Ubuntu/Windows matrix, the package checks, Bun, Deno and
  LibreOffice, so matrix changes never break the ruleset); tag ruleset protecting release tags.
- [ ] The repository must be **public** for provenance. Before going public: secret scan of the
  whole history, commit author emails (rewrite to your noreply address if needed), stale remote
  branches deleted. After: private vulnerability reporting, secret scanning with push protection,
  Dependabot alerts and security updates, rulesets. Commands: `references/repo-hardening.md` of
  the npm-release skill.
- [ ] Never add an `NPM_TOKEN` secret: publishing uses OIDC only.

## A bad release

`npm deprecate turboword@<version> "<reason>"`, then release a fix through the normal flow.
Unpublishing is only possible within 72 hours and breaks anyone who installed that version.
