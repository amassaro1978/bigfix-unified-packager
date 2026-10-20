import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "Invoke-BigFixPackager.ps1").read_text(encoding="utf-8-sig")


def function_body(name: str) -> str:
    start = SOURCE.index(f"function {name}")
    next_function = SOURCE.find("\nfunction ", start + 1)
    return SOURCE[start:] if next_function == -1 else SOURCE[start:next_function]


class PackagerSourceTests(unittest.TestCase):
    def test_parser_detects_msi_before_exe_from_files_folder(self):
        body = function_body("Parse-PsadtFolder")
        self.assertIn('Join-Path $FolderPath "Files"', body)
        self.assertLess(body.index('-Filter "*.msi"'), body.index('-Filter "*.exe"', body.index("$filesFolder")))
        self.assertIn('$result.InstallerType = "MSI"', body)
        self.assertIn('$result.InstallerType = "EXE"', body)
        installer_enumerations = re.findall(
            r'Get-ChildItem -LiteralPath \$filesFolder -Filter "\*\.(?:msi|exe)"[^\r\n]+',
            body,
        )
        self.assertEqual(2, len(installer_enumerations))
        self.assertTrue(all("-ErrorAction Stop" in line for line in installer_enumerations))

    def test_generation_rescans_package_instead_of_trusting_browse_cache(self):
        match = re.search(
            r"\$btnGenScript\.Add_Click\(\{(?P<body>.*?)\n\}\)\n\n# Sign Script",
            SOURCE,
            re.DOTALL,
        )
        self.assertIsNotNone(match)
        body = match.group("body")
        parse_at = body.index("Parse-PsadtFolder -FolderPath $psadtFolder")
        generate_at = body.index("Generate-DeployScript")
        self.assertLess(parse_at, generate_at)
        self.assertIn("-InstallerType $installerType", body)
        self.assertIn("-InstallerFile $installerFile", body)
        self.assertNotIn("-InstallerType $script:InstallerType", body)
        self.assertNotIn("-InstallerFile $script:InstallerFile", body)

    def test_generator_keeps_automatic_msi_and_exe_commands(self):
        body = function_body("Generate-DeployScript")
        self.assertIn('if ($InstallerType -eq "MSI" -and $InstallerFile)', body)
        self.assertIn("Start-ADTMsiProcess -Action Install", body)
        self.assertIn("Start-ADTMsiProcess -Action Uninstall", body)
        self.assertIn('elseif ($InstallerType -eq "EXE" -and $InstallerFile)', body)
        self.assertIn("Start-ADTProcess -FilePath `$exePath", body)


if __name__ == "__main__":
    unittest.main()
