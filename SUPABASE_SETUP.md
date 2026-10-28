# Supabase Auth setup

The login screen uses Supabase email/password authentication. No Supabase credentials
are committed to this repository.

## Create the project and account

1. Create a Supabase project.
2. In **Authentication → Providers → Email**, enable email/password sign-in.
   Disable public sign-ups if accounts should only be provisioned by an administrator.
3. In **Authentication → Users**, create or invite `tenso@bidv.com.vn` and set a
   new password. Do not reuse the old demo password; it was embedded in earlier
   public frontend builds.
4. In **Authentication → URL Configuration**, set the site URL to
   `https://minhhuy2004.github.io/Baocaosolieu12/` and add that URL to the allowed
   redirect URLs.
5. Copy the project URL and the publishable/anon key from **Project Settings → API**.
   Never use or publish the `service_role` key in a browser app.

## Local development

Copy `.env.example` to `.env.local` and replace its values with the project URL
and publishable/anon key. `.env.local` is ignored by Git.

## GitHub Pages deployment

In the repository, open **Settings → Secrets and variables → Actions** and add
these names under either **Variables** (recommended for public client settings)
or **Secrets** to enable sign-in:

- `VITE_SUPABASE_URL`: the Supabase project URL
- `VITE_SUPABASE_ANON_KEY`: the publishable/anon key

The deploy workflow injects these values into the static frontend when they are
configured. They are not required to build or deploy the site; without them,
the login screen explains that Supabase is not configured and sign-in is
unavailable. These values are public client configuration, not server secrets.
Supabase authorization must be enforced by Supabase policies.

## Data protection scope

This integration replaces the hard-coded demo login with real Supabase Auth. The
current report data is still bundled in `src/initialData.json` and client-side code,
so it is downloadable by anyone who can access the public site. Authentication
does not make that bundled data private. Sensitive report data must be moved to
Supabase tables or a protected API with appropriate row-level security before the
data can be considered private.
