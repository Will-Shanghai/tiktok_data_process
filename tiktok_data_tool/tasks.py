from __future__ import annotations

import csv
import os
import runpy
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TASK_CONFIG_FILE = PROJECT_ROOT / "config" / "tool_tasks.csv"
REQUIRED_COLUMNS = {"enabled", "task_id", "task_name", "script_path", "workdir"}
ENABLED_VALUES = {"1", "true", "yes", "y", "启用"}


@dataclass(frozen=True)
class Task:
    task_id: str
    task_name: str
    script_path: str
    workdir: str
    enabled: bool = True

    @property
    def absolute_script_path(self) -> Path:
        return resolve_project_path(self.script_path, "script_path")

    @property
    def absolute_workdir(self) -> Path:
        return resolve_project_path(self.workdir or ".", "workdir")


def resolve_project_path(value: str, field_name: str) -> Path:
    path = (PROJECT_ROOT / value).resolve()
    try:
        path.relative_to(PROJECT_ROOT)
    except ValueError as exc:
        raise ValueError(f"{field_name} 不能指向项目外部路径: {value}") from exc
    return path


def is_enabled(value: str) -> bool:
    return str(value).strip().lower() in ENABLED_VALUES


def load_tasks(config_file: Path = TASK_CONFIG_FILE, include_disabled: bool = False) -> list[Task]:
    if not config_file.exists():
        raise FileNotFoundError(f"任务配置文件不存在: {config_file}")

    with config_file.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        fieldnames = set(reader.fieldnames or [])
        missing = REQUIRED_COLUMNS - fieldnames
        if missing:
            raise ValueError(f"任务配置文件缺少列: {', '.join(sorted(missing))}")

        tasks: list[Task] = []
        seen_ids: set[str] = set()
        for row_no, row in enumerate(reader, start=2):
            enabled = is_enabled(row.get("enabled", ""))
            if not enabled and not include_disabled:
                continue

            task_id = (row.get("task_id") or "").strip()
            task_name = (row.get("task_name") or "").strip()
            script_path = (row.get("script_path") or "").strip()
            workdir = (row.get("workdir") or "").strip()

            blank_fields = [
                name
                for name, value in {
                    "task_id": task_id,
                    "task_name": task_name,
                    "script_path": script_path,
                    "workdir": workdir,
                }.items()
                if not value
            ]
            if blank_fields:
                raise ValueError(f"任务配置第 {row_no} 行缺少: {', '.join(blank_fields)}")
            if task_id in seen_ids:
                raise ValueError(f"任务配置中 task_id 重复: {task_id}")

            task = Task(
                task_id=task_id,
                task_name=task_name,
                script_path=script_path,
                workdir=workdir,
                enabled=enabled,
            )
            if not task.absolute_script_path.exists():
                raise FileNotFoundError(f"任务 {task_id} 的脚本不存在: {task.absolute_script_path}")
            if not task.absolute_workdir.exists():
                raise FileNotFoundError(f"任务 {task_id} 的工作目录不存在: {task.absolute_workdir}")

            seen_ids.add(task_id)
            tasks.append(task)

    return tasks


def find_task(tasks: list[Task], task_id: str) -> Task | None:
    task_id = task_id.strip().lower()
    for task in tasks:
        if task.task_id.lower() == task_id:
            return task
    return None


def run_task(task: Task) -> None:
    old_cwd = os.getcwd()
    os.chdir(task.absolute_workdir)
    try:
        print("\n" + "=" * 60)
        print(f"开始执行: {task.task_name} ({task.task_id})")
        print(f"脚本: {task.absolute_script_path}")
        print(f"工作目录: {task.absolute_workdir}")
        print("=" * 60)
        runpy.run_path(str(task.absolute_script_path), run_name="__main__")
    finally:
        os.chdir(old_cwd)
