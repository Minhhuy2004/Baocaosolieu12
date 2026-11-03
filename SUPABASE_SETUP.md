# Supabase Auth setup

The login screen uses Supabase email/password authentication. The project URL
and publishable key are public client configuration; no service-role key or
user passwords are committed to this repository.

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

## App configuration

The app includes this project's public URL and publishable key, so sign-in can
work on GitHub Pages without Actions variables. These values are intended to be
public in a browser app. Never use a `service_role` key in the frontend.

For local development or a different Supabase project, copy `.env.example` to
`.env.local` and set `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`.
`.env.local` is ignored by Git.

GitHub Actions variables with those names are optional overrides; they are not
required to deploy this project.

## Login requirements

The app signs in with email and password. The account must exist in the
Supabase project's **Authentication → Users** list, have a password, and be
allowed to sign in. Email/password authentication is enabled for the project.
If the login form reports invalid credentials, reset or set that user's
password in Supabase; the publishable key is not a user password.

Supabase authorization must be enforced by Supabase policies.

## Data protection scope

This integration replaces the hard-coded demo login with real Supabase Auth. The
current report data is still bundled in `src/initialData.json` and client-side code,
so it is downloadable by anyone who can access the public site. Authentication
does not make that bundled data private. Sensitive report data must be moved to
Supabase tables or a protected API with appropriate row-level security before the
data can be considered private.
