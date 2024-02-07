from io import BytesIO
import base64


def image_to_base64(image):

    buff = BytesIO()
    image.save(buff, format="PNG")
    img_str = base64.b64encode(buff.getvalue())
    img_str = img_str.decode("utf-8")

    return img_str
