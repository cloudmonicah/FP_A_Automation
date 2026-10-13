"""Every `from <project module> import name` must name something that exists.

Many imports sit inside functions, so a missing name only fails when that code
path runs (for example halfway through a Tableau validation). This catches it
up front.
"""

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PACKAGES = ("automations", "shared", "ui", "app", "config", "scripts")


def _python_files():
    for path in ROOT.rglob("*.py"):
        rel = path.relative_to(ROOT)
        if rel.parts[0] in (".venv", "venv", ".venv-build", "build", "dist") or "manual" in rel.parts:
            continue
        yield path


def _module_name(path: Path) -> str:
    name = ".".join(path.relative_to(ROOT).with_suffix("").parts)
    return name.removesuffix(".__init__")


def _top_level_names(tree: ast.Module) -> set[str]:
    names = set()
    for node in ast.walk(tree):
        parent_is_module = node in tree.body
        in_block = any(node in ast.walk(b) for b in tree.body if isinstance(b, (ast.If, ast.Try)))
        if not (parent_is_module or in_block):
            continue
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, ast.Assign):
            names.update(t.id for t in node.targets for t in ast.walk(t) if isinstance(t, ast.Name))
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            names.update((a.asname or a.name).split(".")[0] for a in node.names)
    return names


def test_project_imports_resolve():
    files = list(_python_files())
    modules = {_module_name(p): p for p in files}
    trees = {p: ast.parse(p.read_text(encoding="utf-8-sig")) for p in files}
    defined = {}
    problems = []
    for path, tree in trees.items():
        for node in ast.walk(tree):
            if not (isinstance(node, ast.ImportFrom) and node.level == 0 and node.module):
                continue
            if node.module.split(".")[0] not in PACKAGES:
                continue
            target = modules.get(node.module)
            if target is None:
                problems.append(f"{path.relative_to(ROOT)}:{node.lineno} no module {node.module}")
                continue
            names = defined.setdefault(node.module, _top_level_names(trees[target]))
            for alias in node.names:
                if alias.name != "*" and alias.name not in names and f"{node.module}.{alias.name}" not in modules:
                    problems.append(f"{path.relative_to(ROOT)}:{node.lineno} {node.module} has no {alias.name}")
    assert problems == []
