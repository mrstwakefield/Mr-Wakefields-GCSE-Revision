# Leaderboard setup

The leaderboard code is already in the site. It switches on once these steps are done.
Students who never sign in are unaffected.

**What's stored online:** first name and surname initial (for example "Amira K."), level, total XP,
this week's XP and the time of the last upload. No email addresses, answers or anything else.

---

## 1. Paste the security rules (Firebase, 2 minutes)

1. Open `firestore.rules` from this repository and copy all of it.
2. Make two changes in your copy (keep them lower case):
   - `school-domain.example` → your school email domain (the part after the @)
   - `teacher@school-domain.example` → your own school email address
3. In the Firebase console go to **Firestore Database → Rules**, delete what's there, paste, then click **Publish**.

Don't save the edited version back to GitHub. The domain should only live in Firebase.

## 2. Allow the website to sign people in (Firebase, 1 minute)

**Authentication → Settings → Authorised domains → Add domain** and enter:

```
mrstwakefield.github.io
```

## 3. Ask school IT to register the app with Microsoft

Forward this to IT:

> Hi, I'd like to let students sign in to my GCSE revision site with their school Microsoft account, for a class XP leaderboard.
> Could you please register an app in Microsoft Entra ID (Azure AD):
>
> - **Name:** GCSE Revision Leaderboard
> - **Supported account types:** Accounts in any organizational directory (multitenant)
> - **Redirect URI (Web):** `https://gcse-fount-of-wisdom.firebaseapp.com/__/auth/handler`
> - **API permissions:** Microsoft Graph delegated `openid`, `email`, `profile`, `User.Read`, and grant admin consent for our tenant
> - **Certificates & secrets:** create a client secret
>
> Then please send me the **Application (client) ID** and the client secret **Value**.
>
> The site only stores each student's first name, surname initial and XP (in Google Firebase).
> No emails, no other personal data. The sign-in is only used to check they're one of our students.

## 4. Switch on Microsoft sign-in (Firebase, 2 minutes)

When IT reply: **Authentication → Sign-in method → Add new provider → Microsoft → Enable**,
paste the **Application ID** and **Application secret**, then **Save**.

Keep the secret in Firebase only. Don't put it in chat, email it on or commit it to GitHub.

## 5. Test it

Open the site, go to **🏆 Leaderboard**, click **Sign in with Microsoft** and use your school account.
You should appear on the board within a few seconds.

On the teacher page (`?teacher`) the **Class leaderboard** box shows everyone who has signed in,
lets you remove anyone, and has a **📽️ Show this week's top 10** button for the projector.

## Good to know

- XP uploads from each student's browser about every 30 seconds while they revise.
- The rules stop scores going down, stop jumps of more than 3,000 XP per upload, and only let
  each student change their own entry. A determined student could still inflate their score
  slowly, so treat it as a bit of fun and use **Remove** on the teacher page if needed.
- The weekly board resets every Monday.
- Students can remove themselves at any time (**Remove me from the board**).
