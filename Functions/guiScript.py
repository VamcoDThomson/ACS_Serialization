import dearpygui.dearpygui as dpg

dpg.create_context()

with dpg.window(width=1000, height=800, pos=[0,0], no_title_bar=True):
    drvNum = dpg.add_input_text(
        label="Drive Drawing Number?",
        hint='#####, or #####REVX',
        width=150,
        uppercase=True,
        no_spaces=True
    )
    numDrvInButton = dpg.add_button(
        label="Confirm",
        width=100,
        
    )
    snWindow = dpg.add_child_window(width=200, height=400, pos=[450,10])
        

dpg.create_viewport(title='Serialize ACS Drives', width=1000, height=800)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
