class DragAndDropManager:
    def __init__(self, root):
        self.root = root
        self.source = None
        self.target = None
        self.initial_widget = None
        self.release_pattern = None
        self.save_cursor = ""

    def start_drag(self, source, event):
        if self.source is not None:
            # Ignore if a drag-and-drop process is already in progress
            return
        self.source = source
        self.initial_widget = event.widget
        self.release_pattern = f"<B1-ButtonRelease>"
        self.save_cursor = self.initial_widget['cursor'] or ""
        self.initial_widget.bind(self.release_pattern, self.on_release)
        self.initial_widget.bind("<B1-Motion>", self.on_motion)
        self.initial_widget['cursor'] = "hand2"

    def on_motion(self, event):
        x, y = event.x_root, event.y_root
        target_widget = self.initial_widget.winfo_containing(x, y)
        new_target = None
        while target_widget is not None:
            try:
                attr = getattr(target_widget, "dnd_accept", None)
            except AttributeError:
                pass
            else:
                new_target = attr(self.source, event)
                if new_target is not None:
                    break
            target_widget = target_widget.master
        old_target = self.target
        if old_target is new_target:
            if old_target is not None:
                old_target.dnd_motion(self.source, event)
        else:
            if old_target is not None:
                self.target = None
                old_target.dnd_leave(self.source, event)
            if new_target is not None:
                new_target.dnd_enter(self.source, event)
                self.target = new_target

    def on_release(self, event):
        self.finish(event, commit=1)

    def finish(self, event, commit=0):
        target = self.target
        source = self.source
        widget = self.initial_widget
        try:
            widget.unbind(self.release_pattern)
            widget.unbind("<B1-Motion>")
            widget['cursor'] = self.save_cursor
            self.target = self.source = self.initial_widget = None
            if target is not None:
                if commit:
                    target.dnd_commit(source, event)
                else:
                    target.dnd_leave(source, event)
        finally:
            source.dnd_end(target, event)