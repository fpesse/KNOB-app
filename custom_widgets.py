import customtkinter as ctk


class WideHandleSlider(ctk.CTkSlider):
    """CTkSlider derivative that keeps an extra-wide handle while masking the track."""

    def __init__(self, *args, track_width=None, **kwargs):
        self._track_width = track_width
        self._track_masks = None
        super().__init__(*args, **kwargs)

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates=no_color_updates)

        if self._track_width is None:
            return
        if self._orientation.lower() != "vertical":
            return

        track_width = self._apply_widget_scaling(self._track_width)
        current_width = self._current_width
        if track_width <= 0 or track_width >= current_width:
            return

        margin = (current_width - track_width) / 2

        if not self._track_masks:
            left_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_left", width=0)
            right_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_right", width=0)
            self._track_masks = (left_id, right_id)

        bg = self._apply_appearance_mode(self._bg_color)
        self._canvas.coords(self._track_masks[0], 0, 0, margin, self._current_height)
        self._canvas.coords(self._track_masks[1], current_width - margin, 0, current_width, self._current_height)
        for mask in self._track_masks:
            self._canvas.itemconfig(mask, fill=bg, outline=bg)
            self._canvas.tag_lower(mask, "slider_parts")
        self._canvas.tag_raise("slider_parts")
