/**
 * Creates a menu in the spreadsheet to manage events
 * 
 * Must have `create_events()', `display_events()`,`update_delete_events()` functions from the library in the target project
 * 
 * Example:
 * ```
 * function onOpen() {
 *  library_name.onOpen();
 * }
 * function create_events() {
 *  library_name.create_events();
 * }
 * function display_events() {
 *  library_name.display_events();
 * }
 * function update_delete_events() {
 *  library_name.update_delete_events();
 * }
 * 
 * ```
 * 
*/
function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Manage Events')
        .addItem('Create Event', 'create_events')
        .addSeparator()
        .addItem('Display Events', 'display_events')
        .addSeparator()
        .addItem('Update/Delete Event', 'update_delete_events')
        .addSeparator()
        .addToUi();
}
