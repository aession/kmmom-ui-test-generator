# Upload To GitHub

Create a private GitHub repository first. Do not initialize it with README, `.gitignore`, or license if you want to push this prepared package directly.

Then run:

```powershell
cd D:\path\to\kmmom-ui-test-generator
git init
git add .
git commit -m "Add KMMOM UI test generator skill"
git branch -M main
git remote add origin https://github.com/<your-org-or-user>/kmmom-ui-test-generator.git
git push -u origin main
```

Before pushing, verify that the repository does not contain:

- browser profiles;
- execution outputs;
- screenshots;
- passwords;
- personal workspace paths;
- real production data;
- local-only config files.

