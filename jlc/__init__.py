from . import plugin


def run():
    from . import plugin
    import importlib
    importlib.reload(plugin)
    plugin.JlcPlugin().Run()
