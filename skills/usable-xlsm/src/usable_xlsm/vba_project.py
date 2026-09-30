"""Experimental source-only MS-OVBA projects, without an Office runtime.

Only standard modules and empty workbook/sheet document modules are emitted.
No performance cache, signature, UserForm, or third-party reference is copied.
Round-trip extraction is a structural check, never an Excel compile test.
"""

from __future__ import annotations

from pathlib import Path
import re
import struct
import uuid

from oletools.olevba import VBA_Parser

from .planning import read_sources
from .syntax import VbaSyntaxError, check_directory

FREE, END, FAT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD


def compress_vba(data: bytes) -> bytes:
    """MS-OVBA 4096-byte chunks with literal/copy tokens and raw fallback."""
    result = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        payload = bytearray()
        position = 0
        while position < len(chunk):
            flag_position = len(payload)
            payload.append(0)
            for bit in range(8):
                if position == len(chunk):
                    break
                offset_bits = max(4, (position - 1).bit_length())
                length_bits = 16 - offset_bits
                maximum = min((1 << length_bits) + 2, len(chunk) - position)
                best_length, best_distance = 0, 0
                if maximum >= 3:
                    search_end = position
                    while search_end > 0:
                        candidate = chunk.rfind(chunk[position:position + 3], 0, search_end + 2)
                        if candidate < 0 or candidate >= position:
                            break
                        length = 3
                        while length < maximum and chunk[candidate + length] == chunk[position + length]:
                            length += 1
                        if length > best_length:
                            best_length, best_distance = length, position - candidate
                        if best_length == maximum:
                            break
                        search_end = candidate
                if best_length >= 3:
                    token = ((best_distance - 1) << length_bits) | (best_length - 3)
                    payload.extend(struct.pack("<H", token))
                    payload[flag_position] |= 1 << bit
                    position += best_length
                else:
                    payload.append(chunk[position])
                    position += 1
        if len(payload) <= 4096:
            result.extend(struct.pack("<H", 0xB000 | (len(payload) - 1)))
            result.extend(payload)
        elif len(chunk) == 4096:
            result.extend(struct.pack("<H", 0x3FFF))
            result.extend(chunk)
        else:
            raise ValueError("Final VBA chunk is incompressible; use an Excel-created seed project (no padding or truncation is applied)")
    return bytes(result)


def _record(identifier: int, value: bytes = b"") -> bytes:
    return struct.pack("<HI", identifier, len(value)) + value


def _u32(value: int) -> bytes:
    return struct.pack("<I", value)


def _project_property(data: bytes, project_id: str, seed: int) -> str:
    """MS-OVBA obfuscation for required *unprotected* project properties."""
    key = sum(project_id.encode("ascii")) & 0xFF
    previous_plain, previous_cipher, older_cipher = key, seed ^ key, seed ^ 2
    output = bytearray([seed, seed ^ 2, seed ^ key])
    ignored = bytes([7]) * ((seed & 6) // 2)
    for value in ignored + _u32(len(data)) + data:
        encrypted = value ^ ((older_cipher + previous_plain) & 0xFF)
        output.append(encrypted)
        older_cipher, previous_cipher, previous_plain = previous_cipher, encrypted, value
    return output.hex().upper()


def _compound_file(streams: dict[str, bytes]) -> bytes:
    """Write bounded CFB v3 with FAT, MiniFAT and balanced directory trees."""
    entries = [dict(name="Root Entry", kind=5, left=FREE, right=FREE, child=FREE, start=END, size=0, color=1)]
    entries.append(dict(name="VBA", kind=1, left=FREE, right=FREE, child=FREE, start=0, size=0, color=1))
    for path, data in sorted(streams.items()):
        entries.append(dict(name=path.split("/")[-1], kind=2, parent=1 if "/" in path else 0, left=FREE, right=FREE, child=FREE, start=END, size=len(data), color=1, data=data))

    def tree(indices: list[int], depth: int, maximum: int) -> int:
        if not indices:
            return FREE
        middle = len(indices) // 2
        index = indices[middle]
        node = entries[index]
        node["color"] = 0 if depth == maximum and depth > 0 else 1
        node["left"] = tree(indices[:middle], depth + 1, maximum)
        node["right"] = tree(indices[middle + 1:], depth + 1, maximum)
        return index

    for parent in (0, 1):
        children = ([1] if parent == 0 else []) + [i for i, item in enumerate(entries) if item.get("parent") == parent]
        children.sort(key=lambda i: (len(entries[i]["name"]), entries[i]["name"].upper()))
        entries[parent]["child"] = tree(children, 0, len(children).bit_length() - 1)

    sectors: list[bytes] = []
    chains: list[list[int]] = []

    def allocate(data: bytes) -> int:
        if not data:
            return END
        ids = list(range(len(sectors), len(sectors) + (len(data) + 511) // 512))
        chains.append(ids)
        sectors.extend(data[i:i + 512].ljust(512, b"\x00") for i in range(0, len(data), 512))
        return ids[0]

    mini = bytearray()
    minifat: list[int] = []
    for item in entries[2:]:
        data = item["data"]
        if 0 < len(data) < 4096:
            start = len(minifat)
            count = (len(data) + 63) // 64
            item["start"] = start
            minifat.extend(range(start + 1, start + count))
            minifat.append(END)
            mini.extend(data.ljust(count * 64, b"\x00"))
        else:
            item["start"] = allocate(data)
    entries[0]["size"] = len(mini)
    entries[0]["start"] = allocate(bytes(mini))
    minifat_data = b"".join(_u32(value) for value in minifat)
    minifat_count = (len(minifat_data) + 511) // 512
    minifat_start = allocate(minifat_data.ljust(minifat_count * 512, b"\xff"))

    directory = bytearray()
    for item in entries:
        name = (item["name"] + "\x00").encode("utf-16le")
        if len(name) > 64:
            raise ValueError("CFB stream name exceeds 31 UTF-16 characters")
        directory.extend(name.ljust(64, b"\x00"))
        directory.extend(struct.pack("<HBBIII", len(name), item["kind"], item["color"], item["left"], item["right"], item["child"]))
        directory.extend(b"\x00" * 36)  # CLSID, state bits, creation/modification times
        directory.extend(struct.pack("<IQ", item["start"], item["size"]))
    directory_start = allocate(bytes(directory))

    fat_count = 1
    while fat_count != (len(sectors) + fat_count + 127) // 128:
        fat_count = (len(sectors) + fat_count + 127) // 128
    if fat_count > 109:
        raise ValueError("Experimental builder size limit exceeded (109 FAT sectors)")
    fat_ids = list(range(len(sectors), len(sectors) + fat_count))
    fat = [FREE] * (fat_count * 128)
    for chain in chains:
        for a, b in zip(chain, chain[1:]):
            fat[a] = b
        fat[chain[-1]] = END
    for index in fat_ids:
        fat[index] = FAT
    fat_data = b"".join(_u32(value) for value in fat)
    sectors.extend(fat_data[i:i + 512] for i in range(0, len(fat_data), 512))

    header = bytearray(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16)
    header.extend(struct.pack("<HHHHH", 0x003E, 3, 0xFFFE, 9, 6))
    header.extend(b"\x00" * 6)
    header.extend(struct.pack("<IIIIIIIII", 0, fat_count, directory_start, 0, 4096, minifat_start, minifat_count, END, 0))
    header.extend(b"".join(_u32(value) for value in fat_ids + [FREE] * (109 - fat_count)))
    return bytes(header) + b"".join(sectors)


def project_bytes(source: str | Path, *, documents: list[str] | None = None, target_os: str = "windows") -> tuple[bytes, dict[str, str]]:
    sources = read_sources(source)
    if any(Path(name).suffix.lower() != ".bas" for name in sources):
        raise ValueError("Experimental builder accepts .bas standard modules only; use an Excel-created project for classes/UserForms.")
    issues = check_directory(source)
    if issues:
        raise VbaSyntaxError(issues)
    documents = ["ThisWorkbook", "Sheet1"] if documents is None else documents
    names = [Path(name).stem for name in sources] + documents
    reserved = {"dir", "_vba_project"}
    if len({name.casefold() for name in names}) != len(names):
        raise ValueError("Module and document names must be unique ignoring case")
    if any(not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,30}", name) or name.casefold() in reserved for name in names):
        raise ValueError("Use ASCII VBA identifiers of at most 31 characters; dir/_VBA_PROJECT are reserved")
    if target_os not in {"windows", "macos", "linux"}:
        raise ValueError("Unknown target OS")
    module_sources = {}
    for filename, text in sources.items():
        name = Path(filename).stem
        attributes = re.findall(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
        if attributes and attributes != [name]:
            raise ValueError(f"Attribute VB_Name does not match filename: {filename}")
        body = re.sub(r"^Attribute .*\n?", "", text.replace("\r\n", "\n"), flags=re.M)
        module_sources[filename] = f'Attribute VB_Name = "{name}"\n' + body
    for name in documents:
        base = "00020819" if name == documents[0] else "00020820"
        module_sources[f"{name}.cls"] = f'Attribute VB_Name = "{name}"\nAttribute VB_Base = "0{{{base}-0000-0000-C000-000000000046}}"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nOption Explicit\n'

    info = _record(1, _u32(2 if target_os == "macos" else 1))
    info += _record(2, _u32(0x409)) + _record(0x14, _u32(0x409))
    info += _record(3, struct.pack("<H", 932)) + _record(4, b"UsableXlsmProject")
    info += _record(5) + _record(0x40) + _record(6) + _record(0x3D)
    info += _record(7, _u32(0)) + _record(8, _u32(0))
    info += struct.pack("<HIIH", 9, 4, 1, 0) + _record(0x0C) + _record(0x3C)
    info += _record(0x0F, struct.pack("<H", len(module_sources))) + _record(0x13, struct.pack("<H", 0xFFFF))
    streams: dict[str, bytes] = {"VBA/_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0, 1)}
    identifier = "{" + str(uuid.uuid5(uuid.NAMESPACE_OID, repr(module_sources))).upper() + "}"
    project = [f'ID="{identifier}"']
    name_map = bytearray()
    for filename, text in module_sources.items():
        name = Path(filename).stem
        encoded = name.encode("cp932")
        unicode_name = name.encode("utf-16le")
        info += _record(0x19, encoded) + _record(0x47, unicode_name)
        info += _record(0x1A, encoded) + _record(0x32, unicode_name)
        info += _record(0x1C) + _record(0x48)
        info += _record(0x31, _u32(0)) + _record(0x1E, _u32(0)) + _record(0x2C, struct.pack("<H", 0xFFFF))
        info += _record(0x22 if name in documents else 0x21) + _record(0x2B)
        # Strict encoding prevents corrupting Japanese or unsupported characters.
        streams[f"VBA/{name}"] = compress_vba(text.replace("\n", "\r\n").encode("cp932"))
        project.append(f"Document={name}/&H00000000" if name in documents else f"Module={name}")
        name_map.extend(encoded + b"\x00" + unicode_name + b"\x00\x00")
    info += _record(0x10)
    project += ['Name="UsableXlsmProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
                f'CMG="{_project_property(bytes(4), identifier, 7)}"',
                f'DPB="{_project_property(bytes(1), identifier, 14)}"',
                f'GC="{_project_property(bytes([255]), identifier, 21)}"',
                "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", ""]
    streams["PROJECT"] = "\r\n".join(project).encode("cp932")
    streams["PROJECTwm"] = bytes(name_map) + b"\x00\x00"
    streams["VBA/dir"] = compress_vba(info)
    payload = _compound_file(streams)
    parser = VBA_Parser("vbaProject.bin", data=payload, relaxed=False)
    try:
        extracted = {name: code for _, _, name, code in parser.extract_macros()}
    finally:
        parser.close()
    if set(extracted) != set(module_sources):
        raise ValueError("Generated VBA project failed module-manifest round-trip")
    for name, text in module_sources.items():
        if extracted[name].replace("\r\n", "\n") != text:
            raise ValueError(f"Generated VBA project did not preserve source: {name}")
    return payload, module_sources


def write_new(path: str | Path, payload: bytes) -> Path:
    """Never overwrite an input or existing candidate."""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("xb") as stream:
        stream.write(payload)
    return destination
