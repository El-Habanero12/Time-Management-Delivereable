from __future__ import annotations

import queue
import threading
from dataclasses import dataclass
from datetime import date, datetime
from typing import Callable, Iterable, List, Optional, Union

import tkinter as tk
from tkinter import ttk, messagebox


@dataclass
class PromptInput:
    timestamp: datetime
    activity: str
    notes: str
    category: str
    energy: int
    focus: int
    prompt_type: str = "regular"
    new_tasks: List[str] | None = None
    worked_task_ids: List[str] | None = None
    completed_task_ids: List[str] | None = None
    task_minutes: int = 0
    task_effort: int = 3
    task_could_be_faster: bool = False


@dataclass
class PromptResult:
    submitted: bool
    dismissed: bool
    prompt_input: Optional[PromptInput] = None
    action: str = "dismiss"


@dataclass
class CatchUpResult:
    submitted: bool
    dismissed: bool
    hours: int
    activity: str = ""
    notes: str = ""
    category: str = "Other"
    energy: int = 3
    focus: int = 3
    split_entries: bool = False


@dataclass
class TaskManagerResult:
    submitted: bool
    dismissed: bool
    added_tasks: List[str]
    completed_task_ids: List[str]
    updated: List[dict]


@dataclass
class SpendingInput:
    amount: float
    entry_type: str
    payment_method: str
    notes: str


@dataclass
class SpendingResult:
    submitted: bool
    dismissed: bool
    spending_input: Optional[SpendingInput] = None


@dataclass
class ReflectionInput:
    date_for: date
    text: str
    tags: str = ""
    created_at: datetime = datetime.now()


@dataclass
class ReflectionResult:
    submitted: bool
    dismissed: bool
    reflection_input: Optional[ReflectionInput] = None


class TkDialogRunner:
    def __init__(self) -> None:
        self._root_ready = threading.Event()
        self._root: Optional[tk.Tk] = None
        self._thread = threading.Thread(target=self._run, name="tk-dialog-runner", daemon=True)
        self._thread.start()
        self._root_ready.wait(timeout=5)
        if not self._root:
            raise RuntimeError("Failed to initialize Tk root")

    def _run(self) -> None:
        root = tk.Tk()
        root.withdraw()
        self._root = root
        self._root_ready.set()
        root.mainloop()

    def run(
        self,
        dialog_func: Callable[[tk.Tk], Union[PromptResult, CatchUpResult, TaskManagerResult, SpendingResult, ReflectionResult]],
    ) -> Union[PromptResult, CatchUpResult, TaskManagerResult, SpendingResult, ReflectionResult]:
        if not self._root:
            raise RuntimeError("Tk root not available")

        result_queue: "queue.Queue[Union[PromptResult, CatchUpResult, TaskManagerResult, SpendingResult, ReflectionResult]]" = queue.Queue(maxsize=1)

        def _invoke() -> None:
            try:
                result = dialog_func(self._root)  # type: ignore[arg-type]
            except Exception:
                result = PromptResult(submitted=False, dismissed=True)
            result_queue.put(result)

        self._root.after(0, _invoke)
        return result_queue.get()


def _base_dialog(root: tk.Tk, title: str) -> tk.Toplevel:
    win = tk.Toplevel(root)
    win.title(title)
    win.attributes("-topmost", True)
    win.resizable(False, False)
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    return win


def _apply_ui_style(root: tk.Tk) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TLabel", font=("Segoe UI", 10))
    style.configure("TButton", font=("Segoe UI", 10))
    style.configure("TEntry", font=("Segoe UI", 10))
    style.configure("TCombobox", font=("Segoe UI", 10))


def prompt_dialog(
    root: tk.Tk,
    categories: Iterable[str],
    timestamp: datetime,
    suggested_category: Optional[str] = None,
    suggestion_confidence: Optional[float] = None,
    quick_entry: bool = False,
    suggest_fn: Optional[Callable[[str], Optional[object]]] = None,
    tasks: Optional[List[dict]] = None,
) -> PromptResult:
    _apply_ui_style(root)
    cats = list(categories) or ["Other"]
    win = _base_dialog(root, "Hourly Check-In")

    win.minsize(560, 520)
    frame = ttk.Frame(win, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    ttk.Label(frame, text="What are you doing right now?").grid(row=0, column=0, columnspan=3, sticky="w")

    activity_var = tk.StringVar(value="")
    notes_var = tk.StringVar(value="")
    category_var = tk.StringVar(value=suggested_category if suggested_category in cats else cats[0])
    energy_var = tk.IntVar(value=3)
    focus_var = tk.IntVar(value=3)
    user_overrode_category = tk.BooleanVar(value=False)
    task_minutes_var = tk.IntVar(value=0)
    task_effort_var = tk.IntVar(value=3)
    task_faster_var = tk.BooleanVar(value=False)

    row = 1
    ttk.Label(frame, text="Activity").grid(row=row, column=0, sticky="w", pady=(8, 2))
    activity_entry = ttk.Entry(frame, textvariable=activity_var, width=60)
    activity_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(8, 2))

    row += 1
    if not quick_entry:
        ttk.Label(frame, text="Notes").grid(row=row, column=0, sticky="w", pady=2)
        notes_entry = ttk.Entry(frame, textvariable=notes_var, width=60)
        notes_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)

        row += 1
        ttk.Label(frame, text="Category").grid(row=row, column=0, sticky="w", pady=2)
        category_box = ttk.Combobox(frame, textvariable=category_var, values=cats, state="readonly", width=28)
        category_box.grid(row=row, column=1, sticky="w", pady=2)

        conf_label = ttk.Label(frame, text="")
        conf_label.grid(row=row, column=2, sticky="w", padx=(8, 0))

        def _mark_override(_: object = None) -> None:
            user_overrode_category.set(True)

        category_box.bind("<<ComboboxSelected>>", _mark_override)

        row += 1
        ttk.Label(frame, text="Energy").grid(row=row, column=0, sticky="w", pady=(6, 2))
        energy_scale = ttk.Scale(frame, from_=1, to=5, orient="horizontal", command=lambda _: None)
        energy_scale.set(energy_var.get())
        energy_scale.grid(row=row, column=1, sticky="ew", pady=2)
        energy_value = ttk.Label(frame, textvariable=energy_var, width=3)
        energy_value.grid(row=row, column=2, sticky="w")

        def _on_energy(val: str) -> None:
            energy_var.set(int(float(val)))

        energy_scale.configure(command=_on_energy)

        row += 1
        ttk.Label(frame, text="Focus").grid(row=row, column=0, sticky="w", pady=(6, 2))
        focus_scale = ttk.Scale(frame, from_=1, to=5, orient="horizontal", command=lambda _: None)
        focus_scale.set(focus_var.get())
        focus_scale.grid(row=row, column=1, sticky="ew", pady=2)
        focus_value = ttk.Label(frame, textvariable=focus_var, width=3)
        focus_value.grid(row=row, column=2, sticky="w")

        def _on_focus(val: str) -> None:
            focus_var.set(int(float(val)))

        focus_scale.configure(command=_on_focus)

        row += 1
        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 8))

        row += 1
        ttk.Label(frame, text="Tasks (add / work / complete)").grid(row=row, column=0, columnspan=3, sticky="w")

        row += 1
        ttk.Label(frame, text="Add tasks (comma-separated)").grid(row=row, column=0, sticky="w", pady=2)
        new_tasks_var = tk.StringVar(value="")
        new_tasks_entry = ttk.Entry(frame, textvariable=new_tasks_var, width=60)
        new_tasks_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)

        open_tasks = tasks or []
        task_labels = [f"{t.get('id')} | {t.get('title')}" for t in open_tasks]

        row += 1
        ttk.Label(frame, text="Worked on (select)").grid(row=row, column=0, sticky="w", pady=2)
        worked_list = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=4, exportselection=False)
        for label in task_labels:
            worked_list.insert(tk.END, label)
        worked_list.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)

        row += 1
        ttk.Label(frame, text="Completed (select)").grid(row=row, column=0, sticky="w", pady=2)
        done_list = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=4, exportselection=False)
        for label in task_labels:
            done_list.insert(tk.END, label)
        done_list.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)

        row += 1
        ttk.Label(frame, text="Minutes spent").grid(row=row, column=0, sticky="w", pady=(6, 2))
        minutes_entry = ttk.Entry(frame, textvariable=task_minutes_var, width=10)
        minutes_entry.grid(row=row, column=1, sticky="w", pady=2)

        row += 1
        ttk.Label(frame, text="Effort (1-5)").grid(row=row, column=0, sticky="w", pady=2)
        effort_scale = ttk.Scale(frame, from_=1, to=5, orient="horizontal", command=lambda _: None)
        effort_scale.set(task_effort_var.get())
        effort_scale.grid(row=row, column=1, sticky="ew", pady=2)
        effort_value = ttk.Label(frame, textvariable=task_effort_var, width=3)
        effort_value.grid(row=row, column=2, sticky="w")

        def _on_effort(val: str) -> None:
            task_effort_var.set(int(float(val)))

        effort_scale.configure(command=_on_effort)

        row += 1
        faster_check = ttk.Checkbutton(frame, text="Could have been faster if I locked in", variable=task_faster_var)
        faster_check.grid(row=row, column=0, columnspan=3, sticky="w", pady=(2, 4))

    row += 1
    button_frame = ttk.Frame(frame)
    button_frame.grid(row=row, column=0, columnspan=3, sticky="e", pady=(10, 0))

    result: PromptResult = PromptResult(submitted=False, dismissed=True)

    def _submit(prompt_type: str = "regular") -> None:
        nonlocal result
        activity = activity_var.get().strip()
        if not activity:
            activity_entry.focus_set()
            return
        new_tasks_raw: List[str] = []
        worked_ids: List[str] = []
        completed_ids: List[str] = []
        if not quick_entry:
            try:
                new_tasks_raw = [t.strip() for t in new_tasks_var.get().split(",") if t.strip()]
            except Exception:
                new_tasks_raw = []
            for idx in worked_list.curselection():
                try:
                    worked_ids.append(open_tasks[idx].get("id"))
                except Exception:
                    pass
            for idx in done_list.curselection():
                try:
                    completed_ids.append(open_tasks[idx].get("id"))
                except Exception:
                    pass
        result = PromptResult(
            submitted=True,
            dismissed=False,
            action=prompt_type,
            prompt_input=PromptInput(
                timestamp=timestamp,
                activity=activity,
                notes=notes_var.get().strip(),
                category=category_var.get().strip() or "Other",
                energy=energy_var.get(),
                focus=focus_var.get(),
                prompt_type=prompt_type,
                new_tasks=new_tasks_raw,
                worked_task_ids=worked_ids,
                completed_task_ids=completed_ids,
                task_minutes=int(task_minutes_var.get() or 0),
                task_effort=int(task_effort_var.get() or 3),
                task_could_be_faster=bool(task_faster_var.get()),
            ),
        )
        win.destroy()

    def _dismiss() -> None:
        win.destroy()

    def _driving() -> None:
        activity_var.set("Driving / In transit")
        _submit(prompt_type="driving")

    log_btn = ttk.Button(button_frame, text="Log", command=_submit)
    log_btn.grid(row=0, column=0, padx=4)
    default_btn = ttk.Button(button_frame, text="I'm in class/driving", command=_driving)
    default_btn.grid(row=0, column=1, padx=4)
    dismiss_btn = ttk.Button(button_frame, text="Dismiss", command=_dismiss)
    dismiss_btn.grid(row=0, column=2, padx=4)

    win.bind("<Return>", lambda _: _submit())
    win.bind("<Escape>", lambda _: _dismiss())

    activity_entry.focus_set()

    if suggest_fn and not quick_entry:
        def _update_suggestion(_: object = None) -> None:
            if user_overrode_category.get():
                return
            activity = activity_var.get().strip()
            if not activity:
                conf_label.configure(text="")
                return
            try:
                suggestion = suggest_fn(activity)
                category = getattr(suggestion, "category", None) if suggestion else None
                confidence = getattr(suggestion, "confidence", None) if suggestion else None
                if category in cats:
                    category_var.set(category)
                    if confidence is not None:
                        conf_label.configure(text=f"Suggested ({float(confidence):.0%})")
                    else:
                        conf_label.configure(text="Suggested")
                else:
                    conf_label.configure(text="")
            except Exception:
                conf_label.configure(text="")

        activity_entry.bind("<KeyRelease>", _update_suggestion)

    win.update_idletasks()
    win.grab_set()
    root.wait_window(win)
    return result


def catch_up_dialog(root: tk.Tk, hours_missed: int, categories: Iterable[str]) -> CatchUpResult:
    _apply_ui_style(root)
    cats = list(categories) or ["Other"]
    win = _base_dialog(root, "Catch-Up Check-In")

    win.minsize(520, 380)
    frame = ttk.Frame(win, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    ttk.Label(frame, text=f"You missed about {hours_missed} hour(s). What were you doing?").grid(
        row=0, column=0, columnspan=3, sticky="w"
    )

    activity_var = tk.StringVar(value="")
    notes_var = tk.StringVar(value="")
    category_var = tk.StringVar(value=cats[0])
    energy_var = tk.IntVar(value=3)
    focus_var = tk.IntVar(value=3)
    split_var = tk.BooleanVar(value=hours_missed > 1)

    row = 1
    ttk.Label(frame, text="Activity").grid(row=row, column=0, sticky="w", pady=(8, 2))
    activity_entry = ttk.Entry(frame, textvariable=activity_var, width=60)
    activity_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(8, 2))

    row += 1
    ttk.Label(frame, text="Notes").grid(row=row, column=0, sticky="w", pady=2)
    notes_entry = ttk.Entry(frame, textvariable=notes_var, width=60)
    notes_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)

    row += 1
    ttk.Label(frame, text="Category").grid(row=row, column=0, sticky="w", pady=2)
    category_box = ttk.Combobox(frame, textvariable=category_var, values=cats, state="readonly", width=28)
    category_box.grid(row=row, column=1, sticky="w", pady=2)

    row += 1
    split_check = ttk.Checkbutton(frame, text="Split into hourly entries", variable=split_var)
    split_check.grid(row=row, column=0, columnspan=3, sticky="w", pady=(6, 2))

    row += 1
    ttk.Label(frame, text="Energy").grid(row=row, column=0, sticky="w", pady=2)
    energy_scale = ttk.Scale(frame, from_=1, to=5, orient="horizontal", command=lambda _: None)
    energy_scale.set(energy_var.get())
    energy_scale.grid(row=row, column=1, sticky="ew", pady=2)
    energy_value = ttk.Label(frame, textvariable=energy_var, width=3)
    energy_value.grid(row=row, column=2, sticky="w")

    def _on_energy(val: str) -> None:
        energy_var.set(int(float(val)))

    energy_scale.configure(command=_on_energy)

    row += 1
    ttk.Label(frame, text="Focus").grid(row=row, column=0, sticky="w", pady=2)
    focus_scale = ttk.Scale(frame, from_=1, to=5, orient="horizontal", command=lambda _: None)
    focus_scale.set(focus_var.get())
    focus_scale.grid(row=row, column=1, sticky="ew", pady=2)
    focus_value = ttk.Label(frame, textvariable=focus_var, width=3)
    focus_value.grid(row=row, column=2, sticky="w")

    def _on_focus(val: str) -> None:
        focus_var.set(int(float(val)))

    focus_scale.configure(command=_on_focus)

    row += 1
    button_frame = ttk.Frame(frame)
    button_frame.grid(row=row, column=0, columnspan=3, sticky="e", pady=(10, 0))

    result: CatchUpResult = CatchUpResult(submitted=False, dismissed=True, hours=hours_missed)

    def _submit() -> None:
        nonlocal result
        activity = activity_var.get().strip()
        if not activity:
            activity_entry.focus_set()
            return
        result = CatchUpResult(
            submitted=True,
            dismissed=False,
            hours=hours_missed,
            activity=activity,
            notes=notes_var.get().strip(),
            category=category_var.get().strip() or "Other",
            energy=energy_var.get(),
            focus=focus_var.get(),
            split_entries=split_var.get(),
        )
        win.destroy()

    def _dismiss() -> None:
        win.destroy()

    log_btn = ttk.Button(button_frame, text="Log", command=_submit)
    log_btn.grid(row=0, column=0, padx=4)
    dismiss_btn = ttk.Button(button_frame, text="Dismiss", command=_dismiss)
    dismiss_btn.grid(row=0, column=1, padx=4)

    win.bind("<Return>", lambda _: _submit())
    win.bind("<Escape>", lambda _: _dismiss())

    activity_entry.focus_set()
    win.update_idletasks()
    win.grab_set()
    root.wait_window(win)
    return result


def task_manager_dialog(root: tk.Tk, tasks: List[dict]) -> TaskManagerResult:
    _apply_ui_style(root)
    win = _base_dialog(root, "Task Manager")

    win.minsize(560, 420)
    frame = ttk.Frame(win, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    ttk.Label(frame, text="Open Tasks").grid(row=0, column=0, columnspan=3, sticky="w")

    task_labels = [f"{t.get('id')} | {t.get('title')}" for t in tasks]
    listbox = tk.Listbox(frame, selectmode=tk.SINGLE, height=8, exportselection=False)
    for label in task_labels:
        listbox.insert(tk.END, label)
    listbox.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 8))

    ttk.Label(frame, text="Add tasks (comma-separated)").grid(row=2, column=0, sticky="w")
    add_var = tk.StringVar(value="")
    add_entry = ttk.Entry(frame, textvariable=add_var, width=60)
    add_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=2)

    ttk.Label(frame, text="Edit title").grid(row=3, column=0, sticky="w")
    title_var = tk.StringVar(value="")
    title_entry = ttk.Entry(frame, textvariable=title_var, width=60)
    title_entry.grid(row=3, column=1, columnspan=2, sticky="ew", pady=2)

    ttk.Label(frame, text="Edit notes").grid(row=4, column=0, sticky="w")
    notes_var = tk.StringVar(value="")
    notes_entry = ttk.Entry(frame, textvariable=notes_var, width=60)
    notes_entry.grid(row=4, column=1, columnspan=2, sticky="ew", pady=2)

    result = TaskManagerResult(submitted=False, dismissed=True, added_tasks=[], completed_task_ids=[], updated=[])

    def _load_selected(_: object = None) -> None:
        selection = listbox.curselection()
        if not selection:
            title_var.set("")
            notes_var.set("")
            return
        idx = selection[0]
        task = tasks[idx]
        title_var.set(str(task.get("title") or ""))
        notes_var.set(str(task.get("notes") or ""))

    listbox.bind("<<ListboxSelect>>", _load_selected)

    def _complete_selected() -> None:
        selection = listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        task_id = tasks[idx].get("id")
        if task_id:
            result.completed_task_ids.append(str(task_id))
            listbox.delete(idx)
            tasks.pop(idx)
            _load_selected()

    def _save_update() -> None:
        selection = listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        task_id = tasks[idx].get("id")
        if task_id:
            result.updated.append(
                {
                    "id": str(task_id),
                    "title": title_var.get().strip(),
                    "notes": notes_var.get().strip(),
                }
            )

    def _submit() -> None:
        nonlocal result
        added = [t.strip() for t in add_var.get().split(",") if t.strip()]
        result.added_tasks.extend(added)
        _save_update()
        result.submitted = True
        result.dismissed = False
        win.destroy()

    def _dismiss() -> None:
        win.destroy()

    button_frame = ttk.Frame(frame)
    button_frame.grid(row=5, column=0, columnspan=3, sticky="e", pady=(10, 0))

    ttk.Button(button_frame, text="Complete", command=_complete_selected).grid(row=0, column=0, padx=4)
    ttk.Button(button_frame, text="Save Edit", command=_save_update).grid(row=0, column=1, padx=4)
    ttk.Button(button_frame, text="Done", command=_submit).grid(row=0, column=2, padx=4)
    ttk.Button(button_frame, text="Cancel", command=_dismiss).grid(row=0, column=3, padx=4)

    win.bind("<Escape>", lambda _: _dismiss())
    add_entry.focus_set()
    win.update_idletasks()
    win.grab_set()
    root.wait_window(win)
    return result


def spending_dialog(root: tk.Tk) -> SpendingResult:
    _apply_ui_style(root)
    win = _base_dialog(root, "Log Today's Spending")
    win.minsize(420, 260)

    frame = ttk.Frame(win, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    amount_var = tk.StringVar(value="")
    type_var = tk.StringVar(value="Expense")
    method_var = tk.StringVar(value="Card")
    notes_var = tk.StringVar(value="")
    error_var = tk.StringVar(value="")

    ttk.Label(frame, text="Amount").grid(row=0, column=0, sticky="w", pady=(4, 2))
    amount_entry = ttk.Entry(frame, textvariable=amount_var, width=20)
    amount_entry.grid(row=0, column=1, sticky="ew", pady=(4, 2))

    ttk.Label(frame, text="Type").grid(row=1, column=0, sticky="w", pady=2)
    type_box = ttk.Combobox(frame, textvariable=type_var, values=["Expense", "Income"], state="readonly", width=18)
    type_box.grid(row=1, column=1, sticky="w", pady=2)

    ttk.Label(frame, text="Payment Method").grid(row=1, column=0, sticky="w", pady=2)
    method_box = ttk.Combobox(
        frame,
        textvariable=method_var,
        values=["Cash", "Withdraw", "Deposit", "Credit Card"],
        state="readonly",
        width=18,
    )
    method_box.grid(row=1, column=1, sticky="e", pady=2)

    ttk.Label(frame, text="Notes/Description (optional)").grid(row=2, column=0, sticky="w", pady=(6, 2))
    notes_entry = ttk.Entry(frame, textvariable=notes_var, width=40)
    notes_entry.grid(row=2, column=1, sticky="ew", pady=(6, 2))

    error_label = ttk.Label(frame, textvariable=error_var, foreground="red")
    error_label.grid(row=3, column=0, columnspan=2, sticky="w")

    result = SpendingResult(submitted=False, dismissed=True)

    def _submit() -> None:
        nonlocal result
        try:
            amount = float(amount_var.get())
            if amount <= 0:
                raise ValueError()
        except Exception:
            error_var.set("Please enter a valid amount (numeric).")
            amount_entry.focus_set()
            return
        result = SpendingResult(
            submitted=True,
            dismissed=False,
            spending_input=SpendingInput(
                amount=amount,
                entry_type=type_var.get(),
                payment_method=method_var.get(),
                notes=notes_var.get().strip(),
            ),
        )
        win.destroy()

    def _dismiss() -> None:
        win.destroy()

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
    ttk.Button(btn_frame, text="Save", command=_submit).grid(row=0, column=0, padx=4)
    ttk.Button(btn_frame, text="Cancel", command=_dismiss).grid(row=0, column=1, padx=4)

    win.bind("<Return>", lambda _: _submit())
    win.bind("<Escape>", lambda _: _dismiss())
    amount_entry.focus_set()
    win.update_idletasks()
    win.grab_set()
    root.wait_window(win)
    return result


def reflection_dialog(root: tk.Tk, date_for: date) -> ReflectionResult:
    _apply_ui_style(root)
    title = f"Daily Reflection - {date_for.isoformat()}"
    win = _base_dialog(root, title)
    win.minsize(520, 420)

    frame = ttk.Frame(win, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    ttk.Label(frame, text=title).grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Label(frame, text="What happened today?").grid(row=1, column=0, sticky="nw", pady=(10, 4))
    text_box = tk.Text(frame, width=70, height=12, wrap="word")
    text_box.grid(row=1, column=1, sticky="nsew", pady=(10, 4))

    ttk.Label(frame, text="Tags (optional)").grid(row=2, column=0, sticky="w", pady=(6, 2))
    tags_var = tk.StringVar(value="")
    tags_entry = ttk.Entry(frame, textvariable=tags_var, width=50)
    tags_entry.grid(row=2, column=1, sticky="ew", pady=(6, 2))

    result = ReflectionResult(submitted=False, dismissed=True)

    def _submit() -> None:
        nonlocal result
        reflection = ReflectionInput(
            date_for=date_for,
            text=text_box.get("1.0", "end").strip(),
            tags=tags_var.get().strip(),
            created_at=datetime.now(),
        )
        result = ReflectionResult(submitted=True, dismissed=False, reflection_input=reflection)
        win.destroy()

    def _dismiss() -> None:
        win.destroy()

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(btn_frame, text="Save", command=_submit).grid(row=0, column=0, padx=4)
    ttk.Button(btn_frame, text="Cancel", command=_dismiss).grid(row=0, column=1, padx=4)

    win.bind("<Return>", lambda _: _submit())
    win.bind("<Control-Return>", lambda _: _submit())
    win.bind("<Escape>", lambda _: _dismiss())
    text_box.focus_set()
    win.update_idletasks()
    win.grab_set()
    root.wait_window(win)
    return result



def error_dialog(root: tk.Tk, title: str, message: str) -> None:
    _apply_ui_style(root)
    messagebox.showerror(title, message, parent=root)
