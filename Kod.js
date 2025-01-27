/** 
 * Zmienne Globalne:
 * @param {string} SECOND_WEBSITE_NAME - Nazwa drugiego serwera ERP, u Was to będzie
 * @param {string} SUBJECT_TEXT - Tytuł wiadomości po której skrypt będzie miał szukać, to jest MEGA WAŻNE JBC
 * @param {int} DAYS_TO_SEARCH - Wiadomości z ilu dni mają być brane pod uwagę?
 * @param {string} COMMENT_MESSAGE - Treśćkomentarza umieszczanego w SECOND_WEBSITE_NAME
 */
var SECOND_WEBSITE_NAME = "webName";
var SUBJECT_TEXT = "Test SA OSS";
var DAYS_TO_SEARCH = 2; // 1 dzień do tyłu, czyli szukaj tylko dzisiejszego dnia
var COMMENT_MESSAGE = "Delivered."

/**
 * Tworzenie UI karty i funkcji dla przycisków Start/Stop
**/
function createHomepage()
{
    var card = CardService.newCardBuilder();

    var section = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
            .setText("Naciśnij przycisk aby dodać komentarze, dodawaj komentarze automatycznie co godzinę."))
        .addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("Wyślij i monitoruj")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("startMonitoring"))
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
            .addButton(CardService.newTextButton()
                .setText("Przestań monitorować")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("stopMonitoring"))));

    card.addSection(section);
    return card.build();
}

/**
 * Funkcja rozpoczynająca tworzenie komentarzy i podpinająca event do monitorowania skrzynki i wysyłania co godzinę.
**/
function startMonitoring()
{
    //Sprawdź maile i uruchom całą sekwencje
    checkNewEmails();
    //Uruchom trigger cogodzinny, który będzie uruchamiał sekwencje
    createTrigger();

    return createNotification("Wiadomość wysłana, monitorowanie cogodzinne rozpoczęte.");
}

/**
 * Funkcja usuwająca trigger
**/
function stopMonitoring()
{
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    return createNotification("Monitorowanie zostało przerwane.");
}

/**
 * Stwórz komunikaty w karcie/UI
**/
function createNotification(message)
{
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText(message))
        .build();
}

/**
 * Tworzenie triggera, który będzie uruchamiany co X czasu, wtedy skrzynka będzie "przeszukiwana"
**/
function createTrigger()
{
    // DUsuń istniejące żeby nie było wycieków pamięci
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    // Stwórz triggera checkNewEmails
    ScriptApp.newTrigger('checkNewEmails')
        .timeBased()
        .everyHours(1)
        .create();
}

/**
 * Funkcja pomocnicza, która ułatwi dobór wyszukiwanych maili, jeśli chcemy przeszukać np. 7 dni wstecz
**/
function getSearchDate()
{
    var date = new Date();
    date.setDate(date.getDate() - (DAYS_TO_SEARCH - 1));
    date.setHours(0, 0, 0, 0);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/**
 * Funkcja, trigger, który będzie uruchamiany cały czas w celu nadpisania ich w systemie 2
**/
function checkNewEmails()
{
    var searchDate = getSearchDate();
    var searchQuery = 'in:anywhere subject:"' + SUBJECT_TEXT + '" after:' + searchDate;

    var threads = GmailApp.search(searchQuery);

    threads.forEach(function (thread)
    {
        var messages = thread.getMessages();
        messages.forEach(function (message)
        {
            processMessage(message);
        });
        //Oznacz jako przeczytane, ale to generalnie w/e
        thread.markRead();
    });
}

/**
 * Funkcja przeszukująca treść maila w celu znalezienia adresu email klienta
 * @param {string} message Treść wiadomości email
**/
function processMessage(message)
{
    try
    {
        var body = message.getPlainBody();
        var customerEmail = extractCustomerEmail(body);

        if (!customerEmail)
        {
            Logger.log('No email found in message: ' + message.getSubject());
            return;
        }

        var filterDates = getFilterDate();
        var orderData = getOrderData(customerEmail, filterDates[0], filterDates[1]);

        if (!orderData || !orderData[0])
        {
            Logger.log('No order found for email: ' + customerEmail);
            return;
        }

        updateOrder(orderData[0].id);
    } catch (err)
    {
        Logger.log('Error processing message: ' + err);
    }
}

/**
 * Pobiera zakres dat do filtrowania zamówień w API Sellasist
 * Obecnie ustawiona data od 1 grudnia 2024 do dnia dzisiejszego * 
 * @returns {Array} Tablica zawierająca:
 *   - dateFrom: Data początkowa w formacie 'YYYY-MM-DD HH:mm:ss'
 *   - dateTo: Data końcowa (dzisiejsza) w formacie 'YYYY-MM-DD HH:mm:ss'
 */
function getFilterDate()
{
    var currentDate = new Date();
    var pastDate = new Date('2024-12-01');
    var dateFrom = pastDate.toISOString().replace('T', ' ').split('.')[0];
    var dateTo = currentDate.toISOString().replace('T', ' ').split('.')[0];
    return [dateFrom, dateTo];
}

/**
 * Funkcja przeszukująca treść maila w celu znalezienia adresu email klienta
 * @param {string} body Treść wiadomości email, jej body, czyli SAMA treśćbez reply to itp, ciągły tekst
**/
function extractCustomerEmail(body)
{
    var match = body.match(/[eE]mail:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/); //Potem ma szukać po dużym "E"
    return match ? match[1] : null;
}


/**
 * Funkcja pobierająca informacje o zamówniu, bo potrzebujemy skądś wziąć order ID
 * @param {string} email Email klienta
 * @param {string} dateFrom Data filtracji Od
 * @param {string} dateTo Data filtracji Do
**/
function getOrderData(email, dateFrom, dateTo)
{
    var url = "https://" + SECOND_WEBSITE_NAME + ".sellasist.pl/api/v1/orders" +
        "?offset=0" +
        "&limit=50" +
        "&email=" + encodeURIComponent(email) +
        "&date_from=" + encodeURIComponent(dateFrom) +
        "&date_to=" + encodeURIComponent(dateTo);

    var options = {
        method: 'get',
        headers: {
            'apiKey': getSellasistApiKey("SECOND_API_KEY"),
            'accept': 'application/json'
        },
        muteHttpExceptions: true
    };

    try
    {
        var response = UrlFetchApp.fetch(url, options);
        return JSON.parse(response.getContentText());
    } catch (err)
    {
        Logger.log('Error getting order: ' + err);
        return null;
    }
}

/**
 * Funkcja aktualizująca status zamówienia w API Sellasist po ID zamównienia
 * @param {int} orderId ID zamówienia
**/
function updateOrder(orderId)
{
    var url = "https://" + SECOND_WEBSITE_NAME + ".sellasist.pl/api/v1/orders/" + orderId;
    var payload = {
        additional_fields: [{
            field_id: 3,
            field_value: COMMENT_MESSAGE
        }]
    };

    var options = {
        method: 'put',
        headers: {
            'apiKey': getSellasistApiKey("SECOND_API_KEY"),
            'accept': 'application/json',
            'content-type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try
    {
        var response = UrlFetchApp.fetch(url, options);
        Logger.log('Order update status: ' + response.getResponseCode());
    } catch (err)
    {
        Logger.log('Error updating order: ' + err);
    }
}

/**
 * Pomocnicza funkcja do pobierania klucza API z ustawień skryptu
 * @param {string} apiName Nazwa klucza API
**/
function getSellasistApiKey(apiName)
{
    return PropertiesService.getScriptProperties().getProperty(apiName);
}