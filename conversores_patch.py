# --- Helpers de estilo para o MDFileManager (compatíveis com KivyMD 2.x) ---
def _apply_if_has(obj, **kwargs):
    """Aplica kwargs somente se a propriedade existir no objeto (evita AttributeError)."""
    for k, v in kwargs.items():
        if hasattr(obj, k):
            setattr(obj, k, v)

def _get_toolbar_from_manager(fm):
    """Tenta obter a toolbar do MDFileManager em builds diferentes."""
    tb = getattr(fm, "toolbar", None)
    if tb:
        return tb
    ids = getattr(fm, "ids", None)
    try:
        return ids.get("toolbar") if ids else None
    except Exception:
        return None

def _style_filemanager(app, fm):
    """
    Aplica o tema desejado ao MDFileManager. 
    Você pode sobrescrever passando um dicionário app.FILEMANAGER_STYLE.
    """
    style = getattr(app, "FILEMANAGER_STYLE", {}) or {}

    # Defaults (podem ser sobrescritos por FILEMANAGER_STYLE)
    sel_bg = style.get("background_color_selection_button", (0.10, 0.40, 0.80, 1))
    sel_icon = style.get("icon_selection_button", "check")
    sel_icon_color = style.get("icon_color", (1, 0.6, 0, 1))

    _apply_if_has(
        fm,
        preview=style.get("preview", False),
        search=style.get("search", "all"),
        use_access=style.get("use_access", True),
        background_color_selection_button=sel_bg,
        icon_selection_button=sel_icon,
        icon_color=sel_icon_color,
    )

    # Toolbar (quando existir)
    tb = _get_toolbar_from_manager(fm)
    if tb:
        _apply_if_has(
            tb,
            title=style.get("toolbar_title", "Selecione o arquivo"),
            md_bg_color=style.get("toolbar_bg", (0.08, 0.08, 0.10, 1)),
            elevation=style.get("toolbar_elevation", 3),
        )
        # Fonte do título (se exposta)
        label = None
        ids = getattr(tb, "ids", None)
        try:
            label = ids.get("label_title") if ids else None
        except Exception:
            label = None
        if label:
            font_name = style.get("font_name", None)
            font_size_sp = style.get("font_size_sp", None)
            if font_name and hasattr(label, "font_name"):
                label.font_name = font_name
            if font_size_sp and hasattr(label, "font_size"):
                label.font_size = f"{font_size_sp}sp"


def _ensure_filemanager(app):
    """Cria (ou reaproveita) um MDFileManager estilizado sem alterar seu fluxo atual."""
    try:
        # Se já existir, só reaplica estilo (caso o tema tenha mudado)
        if hasattr(app, "file_manager") and app.file_manager:
            try:
                _style_filemanager(app, app.file_manager)
            except Exception:
                pass
            return app.file_manager

        from kivymd.uix.filemanager import MDFileManager
        app.file_manager = MDFileManager(
            exit_manager=lambda *a, **k: app.file_manager.close(),
            select_path=lambda *a, **k: None,
            preview=False,
        )
        # extensões aceitas no diálogo (apenas visual)
        try:
            app.file_manager.ext = [".xlsx", ".pkl"]
        except Exception:
            pass

        # aplica o tema/estilo
        try:
            _style_filemanager(app, app.file_manager)
        except Exception:
            pass

        return app.file_manager
    except Exception:
        # fallback mínimo: garante atributo e retorna
        if not hasattr(app, "file_manager"):
            app.file_manager = None
        return app.file_manager


def apply(app=None):
    """Hook opcional chamado pelo main:
    - Se um app for passado, reaplica o estilo no file_manager existente (se houver).
    - Retorna True para indicar que o patch está carregado.
    Uso no main: from conversores_patch import apply as apply_conversores; apply_conversores(self)
    """
    try:
        if app is not None and hasattr(app, "file_manager") and app.file_manager:
            _style_filemanager(app, app.file_manager)
    except Exception:
        pass
    return True
