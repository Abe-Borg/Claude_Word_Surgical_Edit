"""
Tkinter GUI wrapper for the Phase 1 DOCX CSI Normalizer pipeline.

This is a thin wrapper — no business logic. It imports and calls the same
library functions as the smoke test.
"""
from __future__ import annotations

import json
import os
import platform
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, scrolledtext
from pathlib import Path
from typing import Optional




def _load_prompt_file(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Missing required prompt file: {path}")
    try:
        return path.read_text(encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Failed reading prompt file {path}: {exc}") from exc


class LogRedirector:
    """Redirects writes to a thread-safe queue for GUI consumption."""

    def __init__(self, log_queue: queue.Queue) -> None:
        self._queue = log_queue

    def write(self, text: str) -> None:
        if text.strip():
            self._queue.put(text)

    def flush(self) -> None:
        pass


class PipelineThread(threading.Thread):
    """Runs the full Phase 1 pipeline in a background thread."""

    def __init__(
        self,
        docx_path: str,
        api_key: str,
        output_dir: Optional[str],
        log_queue: queue.Queue,
        result_queue: queue.Queue,
    ) -> None:
        super().__init__(daemon=True)
        self.docx_path = docx_path
        self.api_key = api_key
        self.output_dir = output_dir
        self.log_queue = log_queue
        self.result_queue = result_queue

    def _log(self, msg: str) -> None:
        self.log_queue.put(msg)

    def run(self) -> None:
        try:
            from docx_decomposer import (
                extract_docx,
                build_slim_bundle,
                apply_instructions,
                build_style_registry_dict,
            )
            from llm_classifier import classify_document, compute_coverage

            docx_path = Path(self.docx_path)
            extract_dir = Path(f"{docx_path.stem}_extracted")

            # 1) Extract
            self._log(f"Extracting {docx_path.name}...")
            extract_docx(docx_path, extract_dir)

            # 2) Build slim bundle
            self._log("Building slim bundle...")
            bundle = build_slim_bundle(extract_dir)
            bundle_path = extract_dir / "slim_bundle.json"
            bundle_path.write_text(json.dumps(bundle, indent=2), encoding="utf-8")

            n_paras = len(bundle.get("paragraphs", []))
            self._log(f"Slim bundle: {n_paras} paragraphs")

            # 3) Read prompts
            script_dir = Path(__file__).resolve().parent
            master_prompt = _load_prompt_file(script_dir / "master_prompt.txt")
            run_instruction = _load_prompt_file(script_dir / "run_instruction_prompt.txt")

            # 4) Classify
            self._log("Classifying via LLM...")
            instructions = classify_document(
                slim_bundle=bundle,
                master_prompt=master_prompt,
                run_instruction=run_instruction,
                api_key=self.api_key,
            )

            # Save instructions
            instr_path = extract_dir / "instructions.json"
            instr_path.write_text(json.dumps(instructions, indent=2), encoding="utf-8")
            self._log(f"Instructions saved: {instr_path.name}")

            # 5) Coverage
            coverage, styled, classifiable = compute_coverage(bundle, instructions)
            coverage_msg = f"Coverage: {coverage:.1%} ({styled}/{classifiable})"
            self._log(coverage_msg)
            if coverage < 1.0:
                raise ValueError(f"Coverage must be 100% for classifiable paragraphs; got {coverage_msg}")

            # 6) Apply
            self._log("Applying instructions...")
            # Redirect stdout so apply_instructions prints go to our log
            old_stdout = sys.stdout
            sys.stdout = LogRedirector(self.log_queue)
            try:
                apply_instructions(extract_dir, instructions)
            finally:
                sys.stdout = old_stdout

            # 7) Build registries in memory
            self._log("Building style registry...")
            style_registry = build_style_registry_dict(extract_dir, docx_path.name, instructions)

            self._log("Extracting environment...")
            from arch_env_extractor import extract_arch_template_registry
            template_registry = extract_arch_template_registry(extract_dir, docx_path)

            # 8) Validate both registries before writing
            from phase1_validator import validate_phase1_contracts
            self._log("Validating Phase 1 contracts...")
            validate_phase1_contracts(style_registry, template_registry)

            # 9) Write registries (only reached if validation passes)
            reg_path = extract_dir / "arch_style_registry.json"
            reg_path.write_text(json.dumps(style_registry, indent=2), encoding="utf-8")
            self._log(f"Style registry: {reg_path.name}")

            env_path = extract_dir / "arch_template_registry.json"
            env_path.write_text(json.dumps(template_registry, indent=2), encoding="utf-8")
            self._log(f"Environment registry: {env_path.name}")

            # 10) Copy deliverables to output_dir if specified
            output_dir_path = Path(self.output_dir) if self.output_dir else extract_dir
            if output_dir_path != extract_dir:
                import shutil
                output_dir_path.mkdir(parents=True, exist_ok=True)
                shutil.copy2(reg_path, output_dir_path / reg_path.name)
                self._log(f"Copied {reg_path.name} to {output_dir_path}")
                shutil.copy2(env_path, output_dir_path / env_path.name)
                self._log(f"Copied {env_path.name} to {output_dir_path}")

            self.result_queue.put({
                "success": True,
                "extract_dir": str(extract_dir),
                "output_dir": str(output_dir_path),
                "registry_path": str(output_dir_path / reg_path.name),
                "env_path": str(output_dir_path / env_path.name),
                "coverage": coverage_msg,
            })

        except Exception:
            self._log(f"ERROR:\n{traceback.format_exc()}")
            self.result_queue.put({"success": False})


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("DOCX CSI Normalizer — Phase 1")
        root.minsize(600, 500)
        root.configure(bg="#0D0D0D")

        self.colors = {
            "bg_dark": "#0D0D0D",
            "bg_card": "#1A1A1A",
            "bg_input": "#252525",
            "border": "#333333",
            "text_primary": "#FFFFFF",
            "text_secondary": "#B0B0B0",
            "text_muted": "#707070",
            "accent": "#3B82F6",
            "accent_hover": "#2563EB",
            "success": "#22C55E",
            "error": "#EF4444",
        }
        self.fonts = {
            "title": ("Segoe UI", 24, "bold"),
            "subtitle": ("Segoe UI", 12),
            "section": ("Segoe UI", 10, "bold"),
            "label": ("Segoe UI", 11),
            "button": ("Segoe UI", 11, "bold"),
            "mono": ("Consolas", 11),
        }

        self.log_queue: queue.Queue = queue.Queue()
        self.result_queue: queue.Queue = queue.Queue()
        self._result: Optional[dict] = None

        self._build_ui()
        self._poll_queues()

    def _build_ui(self) -> None:
        pad = {"padx": 16, "pady": 10}

        # --- Header ---
        header = tk.Frame(self.root, bg=self.colors["bg_dark"])
        header.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(
            header,
            text="DOCX CSI NORMALIZER",
            bg=self.colors["bg_dark"],
            fg=self.colors["text_primary"],
            font=self.fonts["title"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text="Phase 1 Pipeline • Dark Professional UI",
            bg=self.colors["bg_dark"],
            fg=self.colors["text_secondary"],
            font=self.fonts["subtitle"],
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        # --- Input card ---
        input_frame = tk.Frame(
            self.root,
            bg=self.colors["bg_card"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        input_frame.pack(fill="x", padx=20, pady=(4, 10))

        tk.Label(
            input_frame,
            text="INPUTS",
            bg=self.colors["bg_card"],
            fg=self.colors["text_muted"],
            font=self.fonts["section"],
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 4))

        # Template row
        tk.Label(
            input_frame,
            text="Template",
            bg=self.colors["bg_card"],
            fg=self.colors["text_secondary"],
            font=self.fonts["label"],
        ).grid(row=1, column=0, sticky="w", **pad)
        self.path_var = tk.StringVar()
        tk.Entry(
            input_frame,
            textvariable=self.path_var,
            width=50,
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["mono"],
        ).grid(row=1, column=1, sticky="ew", **pad)
        tk.Button(
            input_frame,
            text="Browse",
            command=self._browse,
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["label"],
        ).grid(row=1, column=2, **pad)

        # API key row
        tk.Label(
            input_frame,
            text="API Key",
            bg=self.colors["bg_card"],
            fg=self.colors["text_secondary"],
            font=self.fonts["label"],
        ).grid(row=2, column=0, sticky="w", **pad)
        self.key_var = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.key_entry = tk.Entry(
            input_frame,
            textvariable=self.key_var,
            width=50,
            show="*",
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["mono"],
        )
        self.key_entry.grid(row=2, column=1, sticky="ew", **pad)
        self._key_visible = False
        self.key_toggle = tk.Button(
            input_frame,
            text="Show",
            command=self._toggle_key,
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["label"],
        )
        self.key_toggle.grid(row=2, column=2, **pad)

        # Output folder row
        tk.Label(
            input_frame,
            text="Output Folder",
            bg=self.colors["bg_card"],
            fg=self.colors["text_secondary"],
            font=self.fonts["label"],
        ).grid(row=3, column=0, sticky="w", **pad)
        self.output_dir_var = tk.StringVar()
        tk.Entry(
            input_frame,
            textvariable=self.output_dir_var,
            width=50,
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["mono"],
        ).grid(row=3, column=1, sticky="ew", **pad)
        tk.Button(
            input_frame,
            text="Browse",
            command=self._browse_output,
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["label"],
        ).grid(row=3, column=2, **pad)

        input_frame.columnconfigure(1, weight=1)

        # --- Run button ---
        self.run_btn = tk.Button(
            self.root,
            text="Run Phase 1",
            command=self._run,
            font=self.fonts["button"],
            height=2,
            bg=self.colors["accent"],
            fg=self.colors["text_primary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            relief="flat",
        )
        self.run_btn.pack(fill="x", padx=20, pady=(2, 12))

        # --- Log card ---
        log_frame = tk.Frame(
            self.root,
            bg=self.colors["bg_card"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        tk.Label(
            log_frame,
            text="ACTIVITY LOG",
            bg=self.colors["bg_card"],
            fg=self.colors["text_muted"],
            font=self.fonts["section"],
        ).pack(anchor="w", padx=14, pady=(10, 6))

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            state="disabled",
            wrap="word",
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            insertbackground=self.colors["text_primary"],
            relief="flat",
            font=self.fonts["mono"],
        )
        self.log_text.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        # --- Status bar ---
        status_frame = tk.Frame(self.root, bg=self.colors["bg_dark"])
        status_frame.pack(fill="x", padx=20, pady=(0, 4))

        self.status_var = tk.StringVar(value="Ready")
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            anchor="w",
            bg=self.colors["bg_dark"],
            fg=self.colors["text_secondary"],
            font=self.fonts["label"],
        )
        self.status_label.pack(side="left")

        # --- Action row ---
        btn_frame = tk.Frame(self.root, bg=self.colors["bg_dark"])
        btn_frame.pack(fill="x", padx=20, pady=(0, 14))

        self.open_folder_btn = tk.Button(
            btn_frame,
            text="Open Output Folder",
            command=self._open_folder,
            state="disabled",
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            disabledforeground=self.colors["text_muted"],
            relief="flat",
            font=self.fonts["label"],
        )
        self.open_folder_btn.pack(side="left", padx=(0, 10), pady=4)

        self.view_reg_btn = tk.Button(
            btn_frame,
            text="View Style Registry",
            command=self._view_registry,
            state="disabled",
            bg=self.colors["bg_input"],
            fg=self.colors["text_secondary"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_primary"],
            disabledforeground=self.colors["text_muted"],
            relief="flat",
            font=self.fonts["label"],
        )
        self.view_reg_btn.pack(side="left", padx=(0, 10), pady=4)

    def _browse(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.path_var.set(path)
            # Auto-populate output folder to same directory as the selected .docx
            self.output_dir_var.set(str(Path(path).parent))

    def _browse_output(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir_var.set(folder)

    def _toggle_key(self) -> None:
        self._key_visible = not self._key_visible
        self.key_entry.config(show="" if self._key_visible else "*")
        self.key_toggle.config(text="Hide" if self._key_visible else "Show")

    def _run(self) -> None:
        docx_path = self.path_var.get().strip()
        api_key = self.key_var.get().strip()

        if not docx_path:
            self.status_var.set("Error: No template selected")
            self.status_label.config(fg=self.colors["error"])
            return
        if not Path(docx_path).exists():
            self.status_var.set("Error: File not found")
            self.status_label.config(fg=self.colors["error"])
            return
        if not api_key:
            self.status_var.set("Error: No API key")
            self.status_label.config(fg=self.colors["error"])
            return

        # Clear log
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        self.run_btn.config(state="disabled")
        self.open_folder_btn.config(state="disabled")
        self.view_reg_btn.config(state="disabled")
        self.status_var.set("Running...")
        self.status_label.config(fg=self.colors["text_secondary"])
        self._result = None

        thread = PipelineThread(
            docx_path=docx_path,
            api_key=api_key,
            output_dir=self.output_dir_var.get().strip() or None,
            log_queue=self.log_queue,
            result_queue=self.result_queue,
        )
        thread.start()

    def _poll_queues(self) -> None:
        # Drain log queue
        while True:
            try:
                msg = self.log_queue.get_nowait()
                self.log_text.config(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
            except queue.Empty:
                break

        # Check result queue
        try:
            result = self.result_queue.get_nowait()
            self._result = result
            self.run_btn.config(state="normal")
            if result["success"]:
                self.status_var.set("Success — " + result.get("coverage", ""))
                self.status_label.config(fg=self.colors["success"])
                self.open_folder_btn.config(state="normal")
                self.view_reg_btn.config(state="normal")

                self.log_text.config(state="normal")
                self.log_text.insert("end", "\nPhase 1 complete. Both registries ready for Phase 2.\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
            else:
                self.status_var.set("Failed — see log for details")
                self.status_label.config(fg=self.colors["error"])
        except queue.Empty:
            pass

        self.root.after(100, self._poll_queues)

    def _open_folder(self) -> None:
        if not self._result:
            return
        folder = self._result.get("output_dir") or self._result.get("extract_dir", "")
        if not folder:
            return
        _open_path(folder)

    def _view_registry(self) -> None:
        if not self._result:
            return
        reg = self._result.get("registry_path", "")
        if not reg:
            return
        _open_path(reg)


def _open_path(path: str) -> None:
    """Open a file or folder with the OS default handler."""
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
