from pathlib import Path

import pytest

import yaml

from at614_editor.domain.bas_extractor import (
    ParameterEntry,
    dump_schema_yaml,
    extract_schema,
    generate_schemas,
)


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "BAS"


def test_extract_schema_with_explicit_indices() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_MINIMAL.bas")

    assert schema.test_id == "MINIMAL"
    assert schema.module == "Module_TEST_MINIMAL.bas"
    assert [(p.index, p.name, p.description) for p in schema.parameters] == [
        (1, "NomeFile", "Nome del file di test"),
        (2, "SourceAddress", "Source address del messaggio"),
        (3, "Tolleranza", "Tolleranza in unita ingegneristiche"),
    ]


def test_extract_schema_with_auto_increment() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_AUTOINC.bas")

    indices = [p.index for p in schema.parameters]
    names = [p.name for p in schema.parameters]

    assert indices == [1, 2, 3, 4]
    assert names == ["Sincro", "FirmwareVersion", "ReleaseHardware", "JumpSpegniModuli"]
    assert schema.parameters[3].description is None


def test_extract_schema_excludes_emax_sentinel_and_keeps_entries_without_e_prefix() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_SENTINEL.bas")

    assert [(p.index, p.name, p.raw_name) for p in schema.parameters] == [
        (1, "ColonnaX", "eColonnaX"),
        (2, "ColonnaY", "eColonnaY"),
        (3, "NomeFileIngresso", "NomeFileIngresso"),
    ]


def test_extract_schema_handles_cp1252_encoded_comments(tmp_path: Path) -> None:
    fixture = tmp_path / "Module_TEST_ACCENTI.bas"
    raw = (
        "Private Enum eTestParameter\n"
        "    eTolleranza = 1                ' Tolleranza espressa in unità ingegneristiche\n"
        "    eRiferimento = 2               ' Riferimento per l'attività di test\n"
        "End Enum\n"
    )
    fixture.write_bytes(raw.encode("cp1252"))

    schema = extract_schema(fixture)

    assert schema.parameters[0].description == "Tolleranza espressa in unità ingegneristiche"
    assert schema.parameters[1].description == "Riferimento per l'attività di test"


def test_extract_schema_ignores_non_test_parameter_enums() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_MINIMAL.bas")

    parameter_names = [p.name for p in schema.parameters]
    assert "Init" not in parameter_names
    assert "Run" not in parameter_names


def test_extract_schema_rejects_non_module_test_filename(tmp_path: Path) -> None:
    fixture = tmp_path / "NotAModule.bas"
    fixture.write_text("Private Enum eTestParameter\n eFoo = 1\nEnd Enum\n", encoding="utf-8")

    with pytest.raises(ValueError):
        extract_schema(fixture)


def test_dump_schema_yaml_is_parseable_and_preserves_fields() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_MINIMAL.bas")
    text = dump_schema_yaml(schema)
    parsed = yaml.safe_load(text)

    assert parsed["id"] == "MINIMAL"
    assert parsed["module"] == "Module_TEST_MINIMAL.bas"
    assert parsed["parameters"][0] == {
        "index": 1,
        "name": "NomeFile",
        "raw_name": "eNomeFile",
        "description": "Nome del file di test",
        "type": "string",
    }


def test_dump_schema_yaml_quotes_values_with_special_chars() -> None:
    schema = extract_schema(FIXTURES_DIR / "Module_TEST_MINIMAL.bas")
    schema.parameters[0] = ParameterEntry(
        index=1,
        name="WithColon",
        raw_name="eWithColon",
        description="value: with colon and # hash",
    )

    text = dump_schema_yaml(schema)
    parsed = yaml.safe_load(text)

    assert parsed["parameters"][0]["description"] == "value: with colon and # hash"


def test_generate_schemas_writes_yaml_file_per_module(tmp_path: Path) -> None:
    output_dir = tmp_path / "auto"

    generated = generate_schemas(FIXTURES_DIR, output_dir)

    assert {path.name for path in generated} == {
        "MINIMAL.yaml",
        "AUTOINC.yaml",
        "SENTINEL.yaml",
    }
    assert all(path.parent == output_dir for path in generated)

    minimal_yaml = yaml.safe_load((output_dir / "MINIMAL.yaml").read_text(encoding="utf-8"))
    assert minimal_yaml["id"] == "MINIMAL"
    assert len(minimal_yaml["parameters"]) == 3


def test_generate_schemas_skips_modules_without_test_parameter_enum(tmp_path: Path) -> None:
    output_dir = tmp_path / "auto"

    generated = generate_schemas(FIXTURES_DIR, output_dir)

    assert "DISPATCHER.yaml" not in {path.name for path in generated}
    assert not (output_dir / "DISPATCHER.yaml").exists()


def test_extract_schema_ignores_specific_suffix_when_companion_module_exists(tmp_path: Path) -> None:
    base = tmp_path / "Module_TEST_MINIMAL.bas"
    base.write_text(
        "Private Enum eTestParameter\n"
        "    eNomeFile = 1               ' Nome del file di test\n"
        "    eSourceAddress = 2           ' Source address del messaggio\n"
        "    eTolleranza = 3              ' Tolleranza in unita ingegneristiche\n"
        "End Enum\n",
        encoding="utf-8",
    )
    suffixed = tmp_path / "Module_TEST_MINIMAL_C.bas"
    suffixed.write_text(base.read_text(encoding="utf-8"), encoding="utf-8")

    sibling_test_ids = {"MINIMAL", "MINIMAL_C"}
    schema = extract_schema(suffixed, sibling_test_ids=sibling_test_ids)

    assert schema.test_id == "MINIMAL"
    assert schema.module == "Module_TEST_MINIMAL_C.bas"


def test_generate_schemas_uses_base_test_id_for_suffixed_modules(tmp_path: Path) -> None:
    fixture = tmp_path / "Module_TEST_MINIMAL.bas"
    fixture.write_text(
        "Private Enum eTestParameter\n"
        "    eNomeFile = 1               ' Nome del file di test\n"
        "    eSourceAddress = 2           ' Source address del messaggio\n"
        "    eTolleranza = 3              ' Tolleranza in unita ingegneristiche\n"
        "End Enum\n",
        encoding="utf-8",
    )
    suffixed = tmp_path / "Module_TEST_MINIMAL_C.bas"
    suffixed.write_text(fixture.read_text(encoding="utf-8"), encoding="utf-8")

    output_dir = tmp_path / "auto"
    generated = generate_schemas(tmp_path, output_dir)

    assert "MINIMAL.yaml" in {path.name for path in generated}
    assert (output_dir / "MINIMAL.yaml").exists()
