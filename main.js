Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
        console.log("Credit Card Sanitizer add-in initialized.");
    }
});

function onItemChanged(eventArgs) {
    Office.context.mailbox.item.getBodyAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let emailBody = result.value;
            const sanitizedBody = sanitizeCreditCardNumbers(emailBody);
            if (sanitizedBody !== emailBody) {
                Office.context.mailbox.item.body.setAsync(
                    sanitizedBody,
                    { coercionType: Office.CoercionType.Html },
                    (setResult) => {
                        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Email body sanitized successfully.");
                        } else {
                            console.error("Failed to set sanitized body:", setResult.error);
                        }
                    }
                );
            }
        } else {
            console.error("Failed to get email body:", result.error);
        }
    });
}

function sanitizeCreditCardNumbers(text) {
    // Regular expression for common credit card formats (Visa, MasterCard, Amex, etc.)
    const ccRegex = /\b(?:\d[ -]*?){13,16}\b/g;
    const ccNumberPattern = /\b(\d{4})[ -]?(\d{4})[ -]?(\d{4})[ -]?(\d{1,4})\b/g;

    // First, verify that the numbers match valid credit card patterns
    let matches = text.match(ccRegex);
    if (!matches) return text;

    // Replace valid credit card numbers, keeping last 4 digits
    return text.replace(ccNumberPattern, (match, p1, p2, p3, p4) => {
        if (isValidCreditCard(match.replace(/[ -]/g, ''))) {
            return `XXXX-XXXX-XXXX-${p4}`;
        }
        return match;
    });
}

function isValidCreditCard(number) {
    // Luhn Algorithm to validate credit card numbers
    let sum = 0;
    let isEven = false;
    for (let i = number.length - 1; i >= 0; i--) {
        let digit = parseInt(number[i]);
        if (isEven) {
            digit *= 2;
            if (digit > 9) digit -= 9;
        }
        sum += digit;
        isEven = !isEven;
    }
    return sum % 10 === 0;
}