# GrowAGarden

## Refresh pets from the wiki

Run the updater to check the Grow a Garden Fandom wiki for pets added after a cutoff date, append any missing pets to `pets_full.json`, rebuild the Markdown sources, and regenerate the main Word document:

```powershell
.\scripts\update_new_pets_since.ps1 -CutoffDate 2026-05-25
```

Useful options:

- `-DryRun` checks the wiki and writes the `new_pets_since_<date>.json` report without changing `pets_full.json`.
- `-SkipGenerate` updates `pets_full.json` but skips rebuilding Markdown and DOCX files.

Primary outputs:

- `pets_full.json`
- `grow_a_garden_pet_guide_by_passive.md`
- `grow_a_garden_pet_guide_alphabetical.md`
- `grow_a_garden_ultimate_pet_guide_alphabetical.docx`
