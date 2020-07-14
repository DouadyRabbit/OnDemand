/**********************************************************************/
/*                                                                    */
/*  ARSOLEEX.H                                                        */
/*                                                                    */
/*  DESCRIPTION:  Header file containing interface definitions for    */
/*                the OnDemand OLE Control.                           */
/*                                                                    */
/* COPYRIGHT:                                                         */
/*  5622-662 (C) COPYRIGHT IBM CORPORATION 1996,2003                  */
/*  All Rights Reserved                                               */
/*  Licensed Materials - Property of IBM                              */
/*                                                                    */
/*  US Government Users Restricted Rights - Use, duplication or       */
/*  disclosure restricted by GSA ADP Schedule Contract with IBM Corp. */
/*                                                                    */
/* NOTE: This program sample is provided on an as-is basis.           */
/*       The licensee of the OnDemand product is free to copy, revise,*/
/*       modify, and make derivative works of this program sample     */
/*       as they see fit.                                             */
/*                                                                    */
/**********************************************************************/

#ifndef _ARSOLE_H_
#define _ARSOLE_H_

/**********************************************************************/
/* ARSOLE Control Constants                                           */
/**********************************************************************/
#define ARS_OLE_SCROLL_RANGE          10000

/**********************************************************************/
/* ARSOLE Control Return Codes                                        */
/**********************************************************************/
#define ARS_OLE_RC_SUCCESS                       0
#define ARS_OLE_RC_NO_MEMORY                     1
#define ARS_OLE_RC_SERVER_ERROR                  2
#define ARS_OLE_RC_USER_CANCELLED                3
#define ARS_OLE_RC_INVALID_DIRECTORY             4
#define ARS_OLE_RC_UNAUTHORIZED_OPERATION        5
#define ARS_OLE_RC_NOT_SUPPORTED                 6
#define ARS_OLE_RC_FILE_ERROR                    7
#define ARS_OLE_RC_ALREADY_LOGGED_ON             8
#define ARS_OLE_RC_NOT_LOGGED_ON                 9
#define ARS_OLE_RC_FOLDER_ALREADY_OPEN           10
#define ARS_OLE_RC_FOLDER_NOT_OPEN               11
#define ARS_OLE_RC_UNKNOWN_FOLDER                12
#define ARS_OLE_RC_NO_FOLDERS_AVAILABLE          13
#define ARS_OLE_RC_DOC_NOT_OPEN                  14
#define ARS_OLE_RC_DOC_ALREADY_OPEN              15
#define ARS_OLE_RC_NO_DOC_AVAILABLE              16
#define ARS_OLE_RC_OPEN_DOC_FAILED               17
#define ARS_OLE_RC_DOC_CANNOT_HORZ_SCROLL        18
#define ARS_OLE_RC_INVALID_INDEX                 19
#define ARS_OLE_RC_INVALID_CONTROL_ID            20
#define ARS_OLE_RC_INVALID_FIELD                 21
#define ARS_OLE_RC_INVALID_OPERATOR              22
#define ARS_OLE_RC_INVALID_MESSAGE_MODE          23
#define ARS_OLE_RC_INVALID_ZOOM_PERCENT          24
#define ARS_OLE_RC_INVALID_PAGE_NUMBER           25
#define ARS_OLE_RC_INVALID_ROTATION              26
#define ARS_OLE_RC_INVALID_COLOR                 27
#define ARS_OLE_RC_INVALID_COPIES                28
#define ARS_OLE_RC_INVALID_ORIENTATION           29
#define ARS_OLE_RC_INVALID_PRINTER               30
#define ARS_OLE_RC_INVALID_FIND_TYPE             31
#define ARS_OLE_RC_INVALID_SCROLL_TYPE           32
#define ARS_OLE_RC_ERROR_DURING_PRINT            33
#define ARS_OLE_RC_DOC_NOT_VIEWABLE_OR_PRINTABLE 34 
#define ARS_OLE_RC_PARM_NOT_SPECIFIED            35
#define ARS_OLE_RC_INVALID_APPL_GROUP_NAME       36
#define ARS_OLE_RC_INVALID_APPL_NAME             37
#define ARS_OLE_RC_INVALID_DATE_FIELD            38
#define ARS_OLE_RC_INVALID_INTEGER_FIELD         39
#define ARS_OLE_RC_INVALID_DECIMAL_FIELD         40
#define ARS_OLE_RC_INVALID_APPLGRP_FIELD_TYPE    41
#define ARS_OLE_RC_DATA_CONVERSION_ERROR         42
#define ARS_OLE_RC_INCORRECT_CURRENT_PASSWORD    43
#define ARS_OLE_RC_PASSWORD_TOO_SHORT            44
#define ARS_OLE_RC_NEW_PASSWORD_MISMATCH         45
#define ARS_OLE_RC_PASSWORD_EXPIRED              46
#define ARS_OLE_RC_INVALID_USER_PASS_SERVER      47
#define ARS_OLE_RC_INVALID_MENU_DATA             48
#define ARS_OLE_RC_INVALID_SELECTION_MODE        49
#define ARS_OLE_RC_AREA_NOT_SELECTED             50
#define ARS_OLE_RC_INVALID_INTENSITY             51
#define ARS_OLE_RC_DEVICE_CONTEXT_NOT_CREATED    52
#define ARS_OLE_RC_DEVICE_CONTEXT_ERROR          53
#define ARS_OLE_RC_IMAGE_PRINT_FAILURE           54
#define ARS_OLE_RC_SERVER_INIT_FAILED            55
#define ARS_OLE_RC_SERVER_CANNOT_CONNECT         56
#define ARS_OLE_RC_SERVER_DLL_NOT_FOUND          57
#define ARS_OLE_RC_SERVER_TIMEOUT                58
#define ARS_OLE_RC_SERVER_DATA_CHANGED           59
#define ARS_OLE_RC_TOO_MANY_VALUE_CHARS          60
#define ARS_OLE_RC_INVALID_MODE                  61

/**********************************************************************/
/* ARSOLE Search Field Operators                                      */
/**********************************************************************/
#define ARS_OLE_OPR_EQUAL                   1
#define ARS_OLE_OPR_NOT_EQUAL               2
#define ARS_OLE_OPR_LESS_THAN               3
#define ARS_OLE_OPR_LESS_THAN_OR_EQUAL      4
#define ARS_OLE_OPR_GREATER_THAN            5
#define ARS_OLE_OPR_GREATER_THAN_OR_EQUAL   6
#define ARS_OLE_OPR_BETWEEN                 7
#define ARS_OLE_OPR_NOT_BETWEEN             8
#define ARS_OLE_OPR_IN                      9
#define ARS_OLE_OPR_NOT_IN                 10
#define ARS_OLE_OPR_LIKE                   11
#define ARS_OLE_OPR_NOT_LIKE               12

/**********************************************************************/
/* ARSOLE Message Modes                                               */
/**********************************************************************/
#define ARS_OLE_USER_MSG_MODE_SHOW      1
#define ARS_OLE_USER_MSG_MODE_SUPPRESS  2

/**********************************************************************/
/* ARSOLE Colors                                                      */
/**********************************************************************/
#define ARS_OLE_COLOR_WHITE    1
#define ARS_OLE_COLOR_BLACK    2
#define ARS_OLE_COLOR_RED      3
#define ARS_OLE_COLOR_BLUE     4
#define ARS_OLE_COLOR_GREEN    5
#define ARS_OLE_COLOR_YELLOW   6
#define ARS_OLE_COLOR_GREY     7
#define ARS_OLE_COLOR_CYAN     8
#define ARS_OLE_COLOR_MAGENTA  9

/**********************************************************************/
/* ARSOLE Rotations                                                   */
/**********************************************************************/
#define ARS_OLE_ROTATION_0      0
#define ARS_OLE_ROTATION_90    90
#define ARS_OLE_ROTATION_180  180
#define ARS_OLE_ROTATION_270  270

/**********************************************************************/
/* ARSOLE Print Orientations                                          */
/**********************************************************************/
#define ARS_OLE_ORIENTATION_PORTRAIT   1
#define ARS_OLE_ORIENTATION_LANDSCAPE  2
#define ARS_OLE_ORIENTATION_BEST_FIT   3
#define ARS_OLE_ORIENTATION_ASIS       4

/**********************************************************************/
/* ARSOLE Iniensities                                                 */
/**********************************************************************/
#define ARS_OLE_INTENSITY_NORMAL   1
#define ARS_OLE_INTENSITY_LIGHT    2
#define ARS_OLE_INTENSITY_NONE     3

/**********************************************************************/
/* ARSOLE Find Types                                                  */
/**********************************************************************/
#define ARS_OLE_FIND_FIRST  1
#define ARS_OLE_FIND_PREV   2
#define ARS_OLE_FIND_NEXT   3

/**********************************************************************/
/* ARSOLE Annotation Status                                           */
/**********************************************************************/
#define ARS_OLE_ANNOTATION_YES      1
#define ARS_OLE_ANNOTATION_NO       2
#define ARS_OLE_ANNOTATION_UNKNOWN  3

/**********************************************************************/
/* ARSOLE Server Printer Types                                        */
/**********************************************************************/
#define ARS_OLE_SERVER_PRINTER_PRINT            1
#define ARS_OLE_SERVER_PRINTER_PRINT_WITH_INFO  2
#define ARS_OLE_SERVER_PRINTER_FAX              3

/**********************************************************************/
/* ARSOLE Server Printer Data Indicies                                */
/**********************************************************************/
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_RECEIVER_NAME         1
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_RECEIVER_COMPANY      2
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_RECEIVER_FAX_NUMBER   3
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SENDER_NAME           4
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SENDER_COMPANY        5
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SENDER_TEL_NUMBER     6
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SENDER_FAX_NUMBER     7
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SENDER_COVER_PAGE     8
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_SUBJECT               9
#define ARS_OLE_SERVER_PRINTER_DATA_FAX_NOTES                10
#define ARS_OLE_SERVER_PRINTER_DATA_INFO_FROM                11
#define ARS_OLE_SERVER_PRINTER_DATA_INFO_TO                  12

/**********************************************************************/
/* ARSOLE Scroll Types                                                */
/**********************************************************************/
#define ARS_OLE_SCROLL_LINEUP           0
#define ARS_OLE_SCROLL_LINELEFT         0
#define ARS_OLE_SCROLL_LINEDOWN         1
#define ARS_OLE_SCROLL_LINERIGHT        1
#define ARS_OLE_SCROLL_PAGEUP           2
#define ARS_OLE_SCROLL_PAGELEFT         2
#define ARS_OLE_SCROLL_PAGEDOWN         3
#define ARS_OLE_SCROLL_PAGERIGHT        3
#define ARS_OLE_SCROLL_THUMBPOSITION    4
#define ARS_OLE_SCROLL_THUMBTRACK       5
#define ARS_OLE_SCROLL_TOP              6
#define ARS_OLE_SCROLL_LEFT             6
#define ARS_OLE_SCROLL_BOTTOM           7
#define ARS_OLE_SCROLL_RIGHT            7
#define ARS_OLE_SCROLL_ENDSCROLL        8

/**********************************************************************/
/* ARSOLE Selection Modes                                             */
/**********************************************************************/
#define ARS_OLE_SELECTION_MODE_NONE      1
#define ARS_OLE_SELECTION_MODE_AREA      2
#define ARS_OLE_SELECTION_MODE_TEXT      3

/**********************************************************************/
/* ARSOLE Document Types                                              */
/**********************************************************************/
#define ARS_OLE_DOC_TYPE_NONE       1
#define ARS_OLE_DOC_TYPE_AFP        2
#define ARS_OLE_DOC_TYPE_BMP        3
#define ARS_OLE_DOC_TYPE_EMAIL      4
#define ARS_OLE_DOC_TYPE_JFIF       5
#define ARS_OLE_DOC_TYPE_GIF        6
#define ARS_OLE_DOC_TYPE_LINE       7
#define ARS_OLE_DOC_TYPE_META       8
#define ARS_OLE_DOC_TYPE_ODDOC      9
#define ARS_OLE_DOC_TYPE_PCX       10
#define ARS_OLE_DOC_TYPE_PDF       11
#define ARS_OLE_DOC_TYPE_SCS       12
#define ARS_OLE_DOC_TYPE_SCS_EXT   13
#define ARS_OLE_DOC_TYPE_TIFF      14
#define ARS_OLE_DOC_TYPE_USER_DEF  15
#define ARS_OLE_DOC_TYPE_PNG       16
#define ARS_OLE_DOC_TYPE_XML       17
/**********************************************************************/
/* ARSOLE Menu Limits                                                 */
/**********************************************************************/
#define ARS_OLE_MAX_MENU_ITEMS       1000
#define ARS_OLE_MAX_MENU_ITEM_LEN    1000
#define ARS_OLE_MIN_MENU_ID            10
#define ARS_OLE_MAX_MENU_ID         65000
#define ARS_OLE_MENU_ID_POPUP           0
#define ARS_OLE_MENU_ID_SEPARATOR       1

/**********************************************************************/
/* ARSOLE Resource Modes                                              */
/**********************************************************************/
#define ARS_OLE_RES_MODE_RETAIN                1
#define ARS_OLE_RES_MODE_ERASE_AFTER_RETRIEVE  2

/**********************************************************************/
/* ARSOLE Types                                                       */
/**********************************************************************/
typedef char ArsOleName[100];

typedef char ArsOleValue[255];

#endif
