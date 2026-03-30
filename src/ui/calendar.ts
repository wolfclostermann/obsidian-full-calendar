/**
 * Handles rendering the calendar given a container element, eventSources, and interaction callbacks.
 */
import {
    Calendar,
    EventApi,
    EventClickArg,
    EventHoveringArg,
    EventSourceInput,
} from "@fullcalendar/core";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import rrulePlugin from "@fullcalendar/rrule";
import listPlugin from "@fullcalendar/list";
import interactionPlugin from "@fullcalendar/interaction";
import googleCalendarPlugin from "@fullcalendar/google-calendar";
import iCalendarPlugin from "@fullcalendar/icalendar";

// There is an issue with FullCalendar RRule support around DST boundaries which is fixed by this monkeypatch:
// https://github.com/fullcalendar/fullcalendar/issues/5273#issuecomment-1360459342
rrulePlugin.recurringTypes[0].expand = function (errd, fr, de) {
    const hours = errd.rruleSet._dtstart.getHours();
    return errd.rruleSet
        .between(de.toDate(fr.start), de.toDate(fr.end), true)
        .map((d: Date) => {
            return new Date(
                Date.UTC(
                    d.getFullYear(),
                    d.getMonth(),
                    d.getDate(),
                    hours,
                    d.getMinutes()
                )
            );
        });
};

const LIST_VIEW_TYPES = ["listWeek", "listMonth", "listAll"] as const;

let listMenuDocClickListener: ((ev: MouseEvent) => void) | null = null;

function closeListViewMenu(): void {
    document
        .querySelectorAll(".ofc-list-view-menu")
        .forEach((el) => el.remove());
    if (listMenuDocClickListener) {
        document.removeEventListener("click", listMenuDocClickListener, true);
        listMenuDocClickListener = null;
    }
}

function syncListMenuButtonActive(viewType: string): void {
    const isListView = LIST_VIEW_TYPES.some((v) => v === viewType);
    document.querySelectorAll(".fc-listMenu-button").forEach((el) => {
        if (isListView) {
            el.classList.add("fc-button-active");
        } else {
            el.classList.remove("fc-button-active");
        }
    });
}

function openListViewMenu(anchorEl: HTMLElement, calendar: Calendar): void {
    closeListViewMenu();
    const menu = document.createElement("div");
    menu.className = "ofc-list-view-menu";
    menu.setAttribute("role", "menu");

    const currentType = calendar.view?.type;
    const items: { view: (typeof LIST_VIEW_TYPES)[number]; label: string }[] = [
        { view: "listWeek", label: "Week" },
        { view: "listMonth", label: "Month" },
        { view: "listAll", label: "All" },
    ];

    for (const { view, label } of items) {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "ofc-list-view-menu-item";
        row.textContent = label;
        row.setAttribute("role", "menuitem");
        if (currentType === view) {
            row.classList.add("is-active");
        }
        row.addEventListener("click", (e) => {
            e.preventDefault();
            e.stopPropagation();
            calendar.changeView(view);
            closeListViewMenu();
        });
        menu.appendChild(row);
    }

    document.body.appendChild(menu);

    const anchor = anchorEl.getBoundingClientRect();
    const menuWidth = menu.offsetWidth;
    let left = anchor.left;
    if (left + menuWidth > window.innerWidth - 8) {
        left = window.innerWidth - menuWidth - 8;
    }
    if (left < 8) {
        left = 8;
    }
    menu.style.left = `${left}px`;
    menu.style.top = `${anchor.bottom + 2}px`;

    listMenuDocClickListener = (ev: MouseEvent) => {
        const t = ev.target as Node;
        if (menu.contains(t)) {
            return;
        }
        const btn = document.querySelector(".fc-listMenu-button");
        if (btn?.contains(t)) {
            return;
        }
        closeListViewMenu();
    };
    window.setTimeout(() => {
        document.addEventListener("click", listMenuDocClickListener!, true);
    }, 0);
}

function toggleListViewMenu(anchorEl: HTMLElement, calendar: Calendar): void {
    if (document.querySelector(".ofc-list-view-menu")) {
        closeListViewMenu();
        return;
    }
    openListViewMenu(anchorEl, calendar);
}

interface ExtraRenderProps {
    eventClick?: (info: EventClickArg) => void;
    select?: (
        startDate: Date,
        endDate: Date,
        allDay: boolean,
        viewType: string
    ) => Promise<void>;
    modifyEvent?: (event: EventApi, oldEvent: EventApi) => Promise<boolean>;
    eventMouseEnter?: (info: EventHoveringArg) => void;
    firstDay?: number;
    initialView?: { desktop: string; mobile: string };
    timeFormat24h?: boolean;
    openContextMenuForEvent?: (
        event: EventApi,
        mouseEvent: MouseEvent
    ) => Promise<void>;
    toggleTask?: (event: EventApi, isComplete: boolean) => Promise<boolean>;
    forceNarrow?: boolean;
}

export function renderCalendar(
    containerEl: HTMLElement,
    eventSources: EventSourceInput[],
    settings?: ExtraRenderProps
): Calendar {
    const isMobile = window.innerWidth < 500;
    const isNarrow = settings?.forceNarrow || isMobile;
    const {
        eventClick,
        select,
        modifyEvent,
        eventMouseEnter,
        openContextMenuForEvent,
        toggleTask,
    } = settings || {};
    const modifyEventCallback =
        modifyEvent &&
        (async ({
            event,
            oldEvent,
            revert,
        }: {
            event: EventApi;
            oldEvent: EventApi;
            revert: () => void;
        }) => {
            const success = await modifyEvent(event, oldEvent);
            if (!success) {
                revert();
            }
        });

    const calendarRef: { current: Calendar | null } = { current: null };

    const cal = new Calendar(containerEl, {
        plugins: [
            // View plugins
            dayGridPlugin,
            timeGridPlugin,
            listPlugin,
            // Drag + drop and editing
            interactionPlugin,
            // Remote sources
            googleCalendarPlugin,
            iCalendarPlugin,
            rrulePlugin,
        ],
        googleCalendarApiKey: "AIzaSyDIiklFwJXaLWuT_4y6I9ZRVVsPuf4xGrk",
        initialView:
            settings?.initialView?.[isNarrow ? "mobile" : "desktop"] ||
            (isNarrow ? "timeGrid3Days" : "timeGridWeek"),
        nowIndicator: true,
        scrollTimeReset: false,
        dayMaxEvents: true,

        headerToolbar: !isNarrow
            ? {
                  left: "prev,next today",
                  center: "title",
                  right: "dayGridMonth,timeGridWeek,timeGridDay,listMenu",
              }
            : !isMobile
            ? {
                  right: "today,prev,next",
                  left: "timeGrid3Days,timeGridDay,listMenu",
              }
            : false,
        footerToolbar: isMobile
            ? {
                  right: "today,prev,next",
                  left: "timeGrid3Days,timeGridDay,listMenu",
              }
            : false,

        customButtons: {
            listMenu: {
                text: isNarrow ? "List" : "List ▾",
                hint: "List view: week, month, or all events",
                click(ev, element) {
                    ev.preventDefault();
                    ev.stopPropagation();
                    const api = calendarRef.current;
                    if (api) {
                        toggleListViewMenu(element, api);
                    }
                },
            },
        },

        views: {
            timeGridDay: {
                type: "timeGrid",
                duration: { days: 1 },
                buttonText: isNarrow ? "1" : "day",
            },
            timeGrid3Days: {
                type: "timeGrid",
                duration: { days: 3 },
                buttonText: "3",
            },
            // FullCalendar's listWeek / listMonth only show events in the visible
            // interval. listAll uses a wide fixed range so every cached event appears.
            // These views are opened from the listMenu toolbar dropdown.
            listWeek: {
                buttonText: "Week",
            },
            listMonth: {
                buttonText: "Month",
            },
            listAll: {
                type: "list",
                buttonText: "All",
                visibleRange: () => ({
                    start: new Date(1990, 0, 1),
                    end: new Date(2060, 11, 31, 23, 59, 59, 999),
                }),
            },
        },

        viewDidMount: (arg) => {
            syncListMenuButtonActive(arg.view.type);
        },
        firstDay: settings?.firstDay,
        ...(settings?.timeFormat24h && {
            eventTimeFormat: {
                hour: "numeric",
                minute: "2-digit",
                hour12: false,
            },
            slotLabelFormat: {
                hour: "numeric",
                minute: "2-digit",
                hour12: false,
            },
        }),
        eventSources,
        eventClick,

        selectable: select && true,
        selectMirror: select && true,
        select:
            select &&
            (async (info) => {
                await select(info.start, info.end, info.allDay, info.view.type);
                info.view.calendar.unselect();
            }),

        editable: modifyEvent && true,
        eventDrop: modifyEventCallback,
        eventResize: modifyEventCallback,

        eventMouseEnter,

        eventDidMount: ({ event, el, textColor }) => {
            el.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                openContextMenuForEvent && openContextMenuForEvent(event, e);
            });
            if (toggleTask) {
                if (event.extendedProps.isTask) {
                    const checkbox = document.createElement("input");
                    checkbox.type = "checkbox";
                    checkbox.checked =
                        event.extendedProps.taskCompleted !== false;
                    checkbox.onclick = async (e) => {
                        e.stopPropagation();
                        if (e.target) {
                            let ret = await toggleTask(
                                event,
                                (e.target as HTMLInputElement).checked
                            );
                            if (!ret) {
                                (e.target as HTMLInputElement).checked = !(
                                    e.target as HTMLInputElement
                                ).checked;
                            }
                        }
                    };
                    // Make the checkbox more visible against different color events.
                    if (textColor == "black") {
                        checkbox.addClass("ofc-checkbox-black");
                    } else {
                        checkbox.addClass("ofc-checkbox-white");
                    }

                    if (checkbox.checked) {
                        el.addClass("ofc-task-completed");
                    }

                    // Depending on the view, we should put the checkbox in a different spot.
                    const container =
                        el.querySelector(".fc-event-time") ||
                        el.querySelector(".fc-event-title") ||
                        el.querySelector(".fc-list-event-title");

                    container?.addClass("ofc-has-checkbox");
                    container?.prepend(checkbox);
                }
            }
        },

        longPressDelay: 250,
    });
    calendarRef.current = cal;
    cal.render();
    syncListMenuButtonActive(cal.view.type);
    return cal;
}
