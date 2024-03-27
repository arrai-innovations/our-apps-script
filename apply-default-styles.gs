const MY_INDENT_FIRST_LINE_MULTIPLIER = 36;
const MY_INDENT_FIRST_LINE_OFFSET = 36;
const MY_INDENT_START_MULTIPLIER = 36;
const MY_INDENT_START_OFFSET = 36;
const MY_INDENT_END = 0;

function onOpen() {
    const ui = DocumentApp.getUi();
    ui.createMenu('Custom Scripts')
        .addItem('Apply Default Styles', 'applyDefaultStylesToSelection')
        .addToUi();
}

function getAllElements(body) {
    const elements = body.getNumChildren();
    const result = [];
    for (let i = 0; i < elements; i++) {
        const element = body.getChild(i);
        result.push(element);
    }
    return result;
}

function applyDefaultStylesToSelection() {
    let i;
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const selection = doc.getSelection();
    const elements = selection?.getRangeElements().map((rangeElement) => rangeElement.getElement()) || getAllElements(body);
    const lists = {};
    for (i = 0; i < elements.length; i++) {
        const element = elements[i];
        switch (element.getType()) {
            case DocumentApp.ElementType.PARAGRAPH:
                const paragraph = element.asParagraph();
                const heading = paragraph.getHeading();
                paragraph.setHeading(heading || DocumentApp.ParagraphHeading.NORMAL);
                break;
            case DocumentApp.ElementType.LIST_ITEM:
                const listItem = element.asListItem();
                const listId = listItem.getListId();
                if (!lists[listId]) {
                    lists[listId] = [];
                }
                lists[listId].push(listItem);
                break;
            default:
                // Do nothing
                break;
        }
    }
    // you cannot remove the last item in a section
    // let's add an empty paragraph to the end of the document, we can trim it later
    const lastP = body.appendParagraph('***last***');
    lastP.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    const lastPText = lastP.getChild(0);
    for (const listId in lists) {
        const listItems = lists[listId];
        const lastListItem = listItems[listItems.length - 1];
        const parent = lastListItem.getParent();
        const lastListItemIndex = parent.getChildIndex(lastListItem);
        const sentinel = parent.insertParagraph(lastListItemIndex + 1, '***SENTINEL***');
        sentinel.setHeading(DocumentApp.ParagraphHeading.NORMAL);
        const sentinelText = sentinel.getChild(0);
        const sentinelAttributes = sentinel.getAttributes();
        const sentinelTextAttributes = sentinelText.getAttributes();
        for (let i = 0; i < listItems.length; i++) {
            const listItem = listItems[i];
            const children = [];
            for (let j = 0; j < listItem.getNumChildren(); j++) {
                children.push(listItem.getChild(j));
            }
            // make a new list item, insert it after sentinel, and copy the children into it
            const sentinelIndex = parent.getChildIndex(sentinel) + i;
            const newListItem = parent.insertListItem(sentinelIndex, '');
            for (let j = 0; j < children.length; j++) {
                const text = children[j].asText();
                const newText = text.copy();
                newText.setAttributes(sentinelTextAttributes);
                newListItem.appendText(newText);
            }
            // order of operations, changing the glyph type before the nesting level causes
            //  the glyph type to change for previous list items at the same nesting level
            newListItem.setNestingLevel(listItem.getNestingLevel());
            newListItem.setGlyphType(listItem.getGlyphType());
            // list styles don't get a default style that is customizable
            // I want list items to act as if NORMAl was applied to them using the UI
            // Only part missing is the indentFirstLine and indentStart, everything else is set by setHeading
            // however setting them here seems to have no effect

            // we avoid this being the last item in the document using lastP.
            listItem.removeFromParent();
        }
        sentinel.removeFromParent();
    }
    for (const listItem of body.getListItems()) {
        // recalculate based on the nesting level.
        // this is hard coded based on my default styles, there is no way to get the default styles...
        // because they are null when default on lastP or lastPText
        const nestingLevel = listItem.getNestingLevel();
        const indentFirstLine = nestingLevel * MY_INDENT_FIRST_LINE_MULTIPLIER + MY_INDENT_FIRST_LINE_OFFSET;
        const indentStart = nestingLevel * MY_INDENT_START_MULTIPLIER + MY_INDENT_START_OFFSET;
        const indentEnd = MY_INDENT_END;
        listItem.setIndentFirstLine(indentFirstLine);
        listItem.setIndentStart(indentStart);
        listItem.setIndentEnd(indentEnd);
        // also, my normal style has space after...
        // there is no way to get the default styles...
        listItem.setSpacingAfter(12);
    }
    // trim the last paragraph
    lastP.clear();
    try {
        lastP.merge();
    } catch (e) {
        // if the preceding node is not a paragraph, we can't merge
        if (e.message !== 'Element must be preceded by an element of the same type.') {
            throw e;
        }
    }
}
