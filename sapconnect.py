import win32com.client


def get_client():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not type(sap_gui_auto) == win32com.client.CDispatch:
        return

    application = sap_gui_auto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sap_gui_auto = None
        return

    for conn in range(application.Children.Count):
        # Loop through the application and get the connection
        connection = application.Children(conn)

        for sess in range(connection.Children.Count):
            session = connection.Children(sess)

            if session.Info.Transaction == 'SESSION_MANAGER':
                print(session)
                return session
            else:
                # Return None and break
                return
