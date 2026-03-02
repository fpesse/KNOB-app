import os
import customtkinter as ctk

try:
    from PIL import Image, ImageTk
except Exception:
    Image = ImageTk = None


class WideHandleSlider(ctk.CTkSlider):
    """CTkSlider derivative that keeps an extra-wide handle while masking the track."""

    def __init__(self, *args, track_width=None, track_end_margin=None, handle_chamfer=None,
                 handle_image_path=None, **kwargs):
        self._track_width = track_width
        self._track_end_margin = track_end_margin
        self._track_side_masks = None
        self._track_end_masks = None
        self._handle_chamfer = handle_chamfer
        self._handle_chamfer_items = None
        self._handle_image_path = handle_image_path if handle_image_path and os.path.exists(handle_image_path) else None
        self._handle_base_image = None
        self._handle_photo = None
        self._handle_photo_size = None
        self._handle_image_id = None
        if self._handle_image_path and Image is not None:
            try:
                with Image.open(self._handle_image_path) as img:
                    self._handle_base_image = img.convert("RGBA")
            except Exception:
                self._handle_base_image = None
        super().__init__(*args, **kwargs)

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates=no_color_updates)

        if self._track_width is None or self._orientation.lower() != "vertical":
            self._clear_track_masks()
        else:
            self._apply_track_masks()

        self._update_handle_overlay()

    # --- Track masking helpers ---
    def _clear_track_masks(self):
        if self._track_side_masks:
            for item in self._track_side_masks:
                self._canvas.delete(item)
            self._track_side_masks = None
        if self._track_end_masks:
            for item in self._track_end_masks:
                self._canvas.delete(item)
            self._track_end_masks = None

    def _apply_track_masks(self):
        track_width = self._apply_widget_scaling(self._track_width)
        current_width = self._current_width
        if track_width <= 0 or track_width >= current_width:
            self._clear_track_masks()
            return

        margin = (current_width - track_width) / 2
        bg = self._apply_appearance_mode(self._bg_color)
        if isinstance(bg, str) and bg == "transparent":
            try:
                bg = self.master.cget("fg_color")
            except Exception:
                bg = "#000000"
            bg = self._apply_appearance_mode(bg)

        if not self._track_side_masks:
            left_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_left", width=0)
            right_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_right", width=0)
            self._track_side_masks = (left_id, right_id)

        self._canvas.coords(self._track_side_masks[0], 0, 0, margin, self._current_height)
        self._canvas.coords(self._track_side_masks[1], current_width - margin, 0, current_width, self._current_height)
        for mask in self._track_side_masks:
            self._canvas.itemconfig(mask, fill=bg, outline=bg)
            self._canvas.tag_lower(mask, "slider_parts")

        # End masks
        end_margin = self._track_end_margin
        if end_margin is None:
            button_length = getattr(self, "_button_length", 0) or 0
            end_margin = button_length / 2
        end_margin = self._apply_widget_scaling(end_margin)

        if end_margin > 0:
            if not self._track_end_masks:
                top_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_top", width=0)
                bottom_id = self._canvas.create_rectangle(0, 0, 0, 0, tags="track_mask_bottom", width=0)
                self._track_end_masks = (top_id, bottom_id)
            self._canvas.coords(self._track_end_masks[0], 0, 0, current_width, end_margin)
            self._canvas.coords(self._track_end_masks[1], 0, self._current_height - end_margin,
                                current_width, self._current_height)
            for mask in self._track_end_masks:
                self._canvas.itemconfig(mask, fill=bg, outline=bg)
                self._canvas.tag_lower(mask, "slider_parts")
        elif self._track_end_masks:
            for mask in self._track_end_masks:
                self._canvas.coords(mask, 0, 0, 0, 0)

    # --- Handle overlays ---
    def _update_handle_overlay(self):
        if self._handle_image_path and Image is not None and ImageTk is not None:
            if self._apply_handle_image():
                self._clear_handle_chamfer_items()
                return
        self._remove_handle_image()
        self._apply_handle_chamfer()

    def _apply_handle_image(self):
        if self._orientation.lower() != "vertical":
            return False

        if self._handle_base_image is None:
            return False

        canvas_width = self._current_width
        canvas_height = self._current_height
        if canvas_width <= 0 or canvas_height <= 0:
            return False

        button_length = getattr(self, "_button_length", 0) or 0
        button_length = self._apply_widget_scaling(button_length)
        if button_length <= 0:
            return False

        target_size = (max(1, int(canvas_width)), max(1, int(button_length)))
        if self._handle_photo is None or self._handle_photo_size != target_size:
            try:
                resized = self._handle_base_image.resize(target_size, Image.LANCZOS)
            except Exception:
                return False
            self._handle_photo = ImageTk.PhotoImage(resized)
            self._handle_photo_size = target_size

        slider_center = self._compute_slider_center(button_length)

        if not self._handle_image_id:
            self._handle_image_id = self._canvas.create_image(0, 0, image=self._handle_photo,
                                                              anchor="center", tags="handle_image")
        self._canvas.coords(self._handle_image_id, canvas_width / 2.0, slider_center)
        self._canvas.itemconfig(self._handle_image_id, image=self._handle_photo)
        self._canvas.tag_raise(self._handle_image_id)
        return True

    def _remove_handle_image(self):
        if self._handle_image_id:
            self._canvas.delete(self._handle_image_id)
            self._handle_image_id = None

    def _compute_slider_center(self, button_length):
        canvas_height = self._current_height
        corner = self._apply_widget_scaling(self._corner_radius)
        usable_height = canvas_height - 2 * corner - button_length
        if usable_height < 0:
            usable_height = 0
        slider_value = getattr(self, "_value", 0.0)
        slider_center = corner + (button_length / 2.0) + usable_height * (1 - slider_value)
        return max(button_length / 2.0, min(canvas_height - button_length / 2.0, slider_center))

    def _clear_handle_chamfer_items(self):
        if self._handle_chamfer_items:
            for item in self._handle_chamfer_items:
                self._canvas.delete(item)
            self._handle_chamfer_items = None

    def _apply_handle_chamfer(self):
        if not self._handle_chamfer or self._orientation.lower() != "vertical":
            self._clear_handle_chamfer_items()
            return

        chamfer = self._apply_widget_scaling(self._handle_chamfer)
        if chamfer <= 0:
            return

        canvas_width = self._current_width
        canvas_height = self._current_height
        if canvas_width <= 0 or canvas_height <= 0:
            return

        button_length = getattr(self, "_button_length", 0) or 0
        button_length = self._apply_widget_scaling(button_length)
        if button_length <= 0:
            return
        corner = self._apply_widget_scaling(self._corner_radius)
        button_corner = self._apply_widget_scaling(self._button_corner_radius)
        slider_center = self._compute_slider_center(button_length)
        top = slider_center - (button_length / 2.0)
        bottom = slider_center + (button_length / 2.0)
        left = button_corner
        right = canvas_width - button_corner

        chamfer = min(chamfer, (right - left) / 2.0, button_length / 2.0)
        if chamfer <= 0:
            return

        # usar el color de fondo del contenedor para "recortar" visualmente el handle
        bg_color = self._apply_appearance_mode(self._bg_color)
        if isinstance(bg_color, str) and bg_color == "transparent":
            try:
                bg_color = self.master.cget("fg_color")
            except Exception:
                bg_color = "#000000"
            bg_color = self._apply_appearance_mode(bg_color)
        track_color = bg_color

        if not self._handle_chamfer_items:
            self._handle_chamfer_items = (
                self._canvas.create_polygon(0, 0, 0, 0, tags="handle_chamfer", width=0),
                self._canvas.create_polygon(0, 0, 0, 0, tags="handle_chamfer", width=0),
                self._canvas.create_polygon(0, 0, 0, 0, tags="handle_chamfer", width=0),
                self._canvas.create_polygon(0, 0, 0, 0, tags="handle_chamfer", width=0),
            )

        coords = [
            (left, top, left + chamfer, top, left, top + chamfer),
            (right - chamfer, top, right, top, right, top + chamfer),
            (right, bottom - chamfer, right, bottom, right - chamfer, bottom),
            (left, bottom - chamfer, left + chamfer, bottom, left, bottom),
        ]

        for item, points in zip(self._handle_chamfer_items, coords):
            self._canvas.coords(item, points)
            self._canvas.itemconfig(item, fill=track_color, outline=track_color)
            self._canvas.tag_raise(item)
