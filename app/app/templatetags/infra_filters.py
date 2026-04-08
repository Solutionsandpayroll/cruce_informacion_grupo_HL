from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Permite {{ dict|get_item:key }} en templates Django."""
    if isinstance(dictionary, dict):
        return dictionary.get(key, "")
    return ""

@register.filter
def get_item_direct(dictionary, key):
    """Alias para get_item, para compatibilidad."""
    if isinstance(dictionary, dict):
        return dictionary.get(key, "")
    return ""