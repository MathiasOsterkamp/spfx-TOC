import { Text } from '@microsoft/sp-core-library';
import * as strings from 'TableOfContentsWebPartStrings';
import {
  LinkValidationError,
  MAX_SLUG_LENGTH,
  MAX_URL_LENGTH,
  validateLinkTitle,
  validateLinkUrl
} from '../../../anchors';

/** Error code -> translated message ('' = valid, as PropertyFieldCollectionData expects). */
export function getValidationMessage(error: LinkValidationError | undefined): string {
  switch (error) {
    case undefined: return '';
    case 'titleRequired': return strings.ErrorTitleRequired;
    case 'titleTooLong': return Text.format(strings.ErrorTitleTooLong, MAX_SLUG_LENGTH);
    case 'titleInvalidChars': return strings.ErrorTitleInvalidChars;
    case 'urlRequired': return strings.ErrorUrlRequired;
    case 'urlTooLong': return Text.format(strings.ErrorUrlTooLong, MAX_URL_LENGTH);
    case 'urlInvalidChars': return strings.ErrorUrlInvalidChars;
    case 'urlUnsafeScheme': return strings.ErrorUrlUnsafeScheme;
    case 'urlNotAbsolute': return strings.ErrorUrlNotAbsolute;
    case 'urlEmptyAnchor': return strings.ErrorUrlEmptyAnchor;
    case 'urlInvalidAnchor': return strings.ErrorUrlInvalidAnchor;
    default: return strings.ErrorUrlInvalidChars;
  }
}

export const getTitleErrorMessage = (value: string): string => getValidationMessage(validateLinkTitle(value));
export const getUrlErrorMessage = (value: string): string => getValidationMessage(validateLinkUrl(value));
