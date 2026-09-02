from __future__ import annotations

import argparse
import sys

from .tasks import Task, find_task, load_tasks, run_task


def print_tasks(tasks: list[Task]) -> None:
    print("可执行任务:")
    for index, task in enumerate(tasks, start=1):
        print(f"{index}. {task.task_name}  task_id={task.task_id}")


def run_one(task: Task) -> bool:
    try:
        run_task(task)
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 0
        if code == 0:
            return True
        print(f"\n任务失败: {task.task_name}，退出码: {code}")
        return False
    except Exception as exc:
        print(f"\n任务失败: {task.task_name}")
        print(f"原因: {exc}")
        return False

    print(f"\n任务完成: {task.task_name}")
    return True


def run_many(tasks: list[Task]) -> int:
    failed: list[Task] = []
    for task in tasks:
        if not run_one(task):
            failed.append(task)

    if failed:
        print("\n以下任务执行失败:")
        for task in failed:
            print(f"- {task.task_name} ({task.task_id})")
        return 1

    print("\n所有任务执行完成。")
    return 0


def interactive_menu(tasks: list[Task]) -> int:
    while True:
        print("\nTikTok 数据处理工具")
        print("-" * 24)
        print_tasks(tasks)
        print("A. 全部执行")
        print("Q. 退出")

        choice = input("\n请选择任务编号、task_id、A 或 Q: ").strip()
        if not choice:
            continue
        if choice.lower() == "q":
            return 0
        if choice.lower() == "a":
            return run_many(tasks)
        if choice.isdigit():
            index = int(choice)
            if 1 <= index <= len(tasks):
                return 0 if run_one(tasks[index - 1]) else 1
            print("任务编号不存在，请重新选择。")
            continue

        task = find_task(tasks, choice)
        if task:
            return 0 if run_one(task) else 1
        print("未找到对应任务，请重新选择。")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="TikTok 数据处理统一入口")
    parser.add_argument(
        "task",
        nargs="?",
        help="要执行的 task_id；使用 all 执行全部。不填则进入菜单。",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="列出 config/tool_tasks.csv 中启用的任务。",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    tasks = load_tasks()

    if args.list:
        print_tasks(tasks)
        return

    if not tasks:
        print("没有启用的任务，请检查 config/tool_tasks.csv。")
        sys.exit(1)

    if not args.task:
        sys.exit(interactive_menu(tasks))

    task_id = args.task.strip()
    if task_id.lower() == "all":
        sys.exit(run_many(tasks))

    task = find_task(tasks, task_id)
    if not task:
        print(f"未找到任务: {task_id}")
        print_tasks(tasks)
        sys.exit(1)

    sys.exit(0 if run_one(task) else 1)
