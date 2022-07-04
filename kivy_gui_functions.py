from textwrap import wrap


def label_text_wrap(text: str, label_width, label) -> str:
    if label.texture_size[0] == 0:
        return text

    if label.texture_size[0] != label_width:
        max_line_length = len((lines := label.text.split("\n"))[0])
        for line in lines:
            if (new_len := len(line)) > max_line_length:
                max_line_length = new_len

        char_width = label.texture_size[0] / max_line_length
        wrapped_text = '\n'.join(wrap(text, int(label_width / char_width)))
        return wrapped_text
    else:
        return text
