# emailextract.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Email text extraction application."""

if __name__ == "__main__":

    from . import APPLICATION_NAME

    try:
        from solentware_misc.gui.startstop import (
            start_application_exception,
            stop_application,
            application_exception,
        )
    except Exception as error:
        import tkinter.messagebox

        try:
            tkinter.messagebox.showerror(
                title="Start Exception",
                message=".\n\nThe reported exception is:\n\n".join(
                    (
                        " ".join(
                            (
                                "Unable to import",
                                "solentware_misc.gui.startstop module",
                            )
                        ),
                        str(error),
                    )
                ),
            )
        except BaseException:
            pass
        raise SystemExit("Unable to import start application utilities")
    try:
        from .gui.select import Select
    except Exception as error:
        start_application_exception(
            error, appname=APPLICATION_NAME, action="import"
        )
        raise SystemExit(" import ".join(("Unable to", APPLICATION_NAME)))
    try:
        app = Select(title=APPLICATION_NAME, width=400, height=200)
    except Exception as error:
        start_application_exception(
            error, appname=APPLICATION_NAME, action="initialise"
        )
        raise SystemExit(" initialise ".join(("Unable to", APPLICATION_NAME)))
    try:
        app.root.mainloop()
    except SystemExit:
        stop_application(app, app.root)
        raise
    except Exception as error:
        application_exception(
            error,
            app,
            app.root,
            title=APPLICATION_NAME,
            appname=APPLICATION_NAME,
        )
