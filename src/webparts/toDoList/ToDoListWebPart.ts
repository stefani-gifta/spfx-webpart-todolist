import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ToDoListWebPart.module.scss';
import * as strings from 'ToDoListWebPartStrings';

export interface IToDoListWebPartProps {
  description: string;
}

type TaskStatus = 'Not Started' | 'In Progress' | 'Completed';

interface ITask {
  id: string;
  title: string;
  status: TaskStatus;
  createdAt: number;
}

const STATUS_VALUES: TaskStatus[] = ['Not Started', 'In Progress', 'Completed'];

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {

  private _tasks: ITask[] = [];
  private _isCollapsed: boolean = false;
  private _editingTaskId: string | undefined = undefined;
  private _eventsBound: boolean = false;
  private _storageKey: string = '';

  protected onInit(): Promise<void> {
    this._storageKey = `spfx-todolist-${this.context.instanceId}`;
    this._loadTasks();
    return Promise.resolve();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.toDoList}">
        <div class="${styles.card}">
          <button type="button" class="${styles.header}" data-action="toggle-collapse" aria-expanded="${!this._isCollapsed}">
            <span class="${styles.headerLeft}">
              <svg class="${styles.chevron}" width="12" height="12" viewBox="0 0 12 12" aria-hidden="true" focusable="false">
                <path d="M3 2l5 4-5 4" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
              <span class="${styles.title}">${escape(this.properties.description) || 'To do list'}</span>
            </span>
            <span class="${styles.stats}" data-role="stats"></span>
          </button>
          <div class="${styles.body} ${this._isCollapsed ? styles.collapsed : ''}" data-role="body">
            <form class="${styles.addForm}" data-role="add-form">
              <input
                type="text"
                class="${styles.addInput}"
                data-role="add-input"
                placeholder="Add a task"
                aria-label="New task title"
                maxlength="200"
              />
              <button type="submit" class="${styles.addButton}">Add task</button>
            </form>
            <div data-role="list-container"></div>
          </div>
        </div>
      </section>`;

    this._renderStats();
    this._renderTaskList();

    if (!this._eventsBound) {
      this._bindEvents();
      this._eventsBound = true;
    }
  }

  // ---------- rendering helpers ----------

  private _renderStats(): void {
    const statsEl = this.domElement.querySelector('[data-role="stats"]');
    if (!statsEl) { return; }

    const total = this._tasks.length;
    const completed = this._tasks.filter(t => t.status === 'Completed').length;
    const pending = total - completed;

    statsEl.innerHTML = `
      <span class="${styles.statPill} ${styles.statTotal}">${total} total</span>
      <span class="${styles.statPill} ${styles.statPending}">${pending} pending</span>
      <span class="${styles.statPill} ${styles.statCompleted}">${completed} completed</span>
    `;
  }

  private _renderTaskList(): void {
    const container = this.domElement.querySelector('[data-role="list-container"]');
    if (!container) { return; }

    if (this._tasks.length === 0) {
      container.innerHTML = `<p class="${styles.emptyState}">No tasks yet. Add one above to get started.</p>`;
      return;
    }

    container.innerHTML = `
      <ul class="${styles.taskList}">
        ${this._tasks.map(task => this._renderTaskItem(task)).join('')}
      </ul>
    `;

    if (this._editingTaskId) {
      const input = this.domElement.querySelector('[data-role="edit-input"]') as HTMLInputElement | null;
      if (input) {
        input.focus();
        input.setSelectionRange(input.value.length, input.value.length);
      }
    }
  }

  private _renderTaskItem(task: ITask): string {
    const isEditing = this._editingTaskId === task.id;

    const titleMarkup = isEditing
      ? `<input type="text" class="${styles.taskTitleInput}" data-role="edit-input" data-id="${task.id}" value="${escape(task.title)}" maxlength="200" />`
      : `<span class="${styles.taskTitle}">${escape(task.title)}</span>`;

    const actionsMarkup = isEditing
      ? `
        <button type="button" class="${styles.iconButton}" data-action="save-edit" data-id="${task.id}" aria-label="Save task">✓</button>
        <button type="button" class="${styles.iconButton}" data-action="cancel-edit" data-id="${task.id}" aria-label="Cancel edit">✕</button>
      `
      : `
        <button type="button" class="${styles.iconButton}" data-action="edit" data-id="${task.id}" aria-label="Edit task">✎</button>
        <button type="button" class="${styles.iconButton} ${styles.iconButtonDanger}" data-action="delete" data-id="${task.id}" aria-label="Delete task">🗑</button>
      `;

    return `
      <li class="${styles.taskItem} ${this._statusClass(task.status)}" data-id="${task.id}">
        <div class="${styles.taskMain}">
          ${titleMarkup}
        </div>
        <div class="${styles.taskControls}">
          <select class="${styles.statusSelect}" data-role="status-select" data-id="${task.id}" aria-label="Task status">
            ${STATUS_VALUES.map(s => `<option value="${s}" ${s === task.status ? 'selected' : ''}>${s}</option>`).join('')}
          </select>
          ${actionsMarkup}
        </div>
      </li>
    `;
  }

  private _statusClass(status: TaskStatus): string {
    switch (status) {
      case 'In Progress': return styles.statusInProgress;
      case 'Completed': return styles.statusCompleted;
      default: return styles.statusNotStarted;
    }
  }

  // ---------- events (delegated, bound once on the persistent domElement) ----------

  private _bindEvents(): void {
    this.domElement.addEventListener('click', this._onClick);
    this.domElement.addEventListener('submit', this._onSubmit);
    this.domElement.addEventListener('change', this._onChange);
    this.domElement.addEventListener('keydown', this._onKeyDown);
  }

  private _onClick = (ev: Event): void => {
    const target = ev.target as HTMLElement;
    const actionEl = target.closest('[data-action]') as HTMLElement | null;
    if (!actionEl) { return; }

    const action = actionEl.getAttribute('data-action');
    const id = actionEl.getAttribute('data-id') || undefined;

    switch (action) {
      case 'toggle-collapse':
        ev.preventDefault();
        this._toggleCollapse();
        break;
      case 'edit':
        this._startEdit(id);
        break;
      case 'save-edit':
        this._commitEdit(id);
        break;
      case 'cancel-edit':
        this._cancelEdit();
        break;
      case 'delete':
        this._deleteTask(id);
        break;
      default:
        break;
    }
  };

  private _onSubmit = (ev: Event): void => {
    const form = (ev.target as HTMLElement).closest('[data-role="add-form"]');
    if (!form) { return; }
    ev.preventDefault();
    this._addTaskFromInput();
  };

  private _onChange = (ev: Event): void => {
    const target = ev.target as HTMLElement;
    if (target.matches && target.matches('[data-role="status-select"]')) {
      const id = target.getAttribute('data-id') || undefined;
      const value = (target as HTMLSelectElement).value as TaskStatus;
      this._updateStatus(id, value);
    }
  };

  private _onKeyDown = (ev: KeyboardEvent): void => {
    const target = ev.target as HTMLElement;
    if (target.matches && target.matches('[data-role="edit-input"]')) {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        this._commitEdit(target.getAttribute('data-id') || undefined);
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        this._cancelEdit();
      }
    }
  };

  // ---------- actions ----------

  private _toggleCollapse(): void {
    this._isCollapsed = !this._isCollapsed;
    const body = this.domElement.querySelector('[data-role="body"]');
    const header = this.domElement.querySelector(`.${styles.header}`);
    if (body) {
      body.classList.toggle(styles.collapsed, this._isCollapsed);
    }
    if (header) {
      header.setAttribute('aria-expanded', String(!this._isCollapsed));
    }
  }

  private _addTaskFromInput(): void {
    const input = this.domElement.querySelector('[data-role="add-input"]') as HTMLInputElement | null;
    if (!input) { return; }

    const title = input.value.trim();
    if (!title) { return; }

    this._tasks.push({
      id: this._generateId(),
      title,
      status: 'Not Started',
      createdAt: Date.now()
    });

    input.value = '';
    this._saveTasks();
    this._renderStats();
    this._renderTaskList();
    input.focus();
  }

  private _startEdit(id: string | undefined): void {
    if (!id) { return; }
    this._editingTaskId = id;
    this._renderTaskList();
  }

  private _commitEdit(id: string | undefined): void {
    if (!id) { return; }
    const input = this.domElement.querySelector('[data-role="edit-input"]') as HTMLInputElement | null;
    const task = this._tasks.filter(t => t.id === id)[0];

    if (task && input) {
      const newTitle = input.value.trim();
      if (newTitle) {
        task.title = newTitle;
      }
    }

    this._editingTaskId = undefined;
    this._saveTasks();
    this._renderTaskList();
  }

  private _cancelEdit(): void {
    this._editingTaskId = undefined;
    this._renderTaskList();
  }

  private _deleteTask(id: string | undefined): void {
    if (!id) { return; }
    const task = this._tasks.filter(t => t.id === id)[0];
    if (!task) { return; }

    const confirmed = window.confirm(`Delete "${task.title}"?`);
    if (!confirmed) { return; }

    this._tasks = this._tasks.filter(t => t.id !== id);
    this._saveTasks();
    this._renderStats();
    this._renderTaskList();
  }

  private _updateStatus(id: string | undefined, status: TaskStatus): void {
    if (!id) { return; }
    const task = this._tasks.filter(t => t.id === id)[0];
    if (!task) { return; }

    task.status = status;
    this._saveTasks();
    this._renderStats();
    this._renderTaskList();
  }

  // ---------- storage ----------

  private _loadTasks(): void {
    try {
      const raw = window.localStorage.getItem(this._storageKey);
      this._tasks = raw ? JSON.parse(raw) as ITask[] : [];
    } catch (e) {
      this._tasks = [];
    }
  }

  private _saveTasks(): void {
    try {
      window.localStorage.setItem(this._storageKey, JSON.stringify(this._tasks));
    } catch (e) {
      // storage unavailable (private browsing / quota) - tasks remain in memory for this session
    }
  }

  private _generateId(): string {
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  }

  // ---------- theme ----------

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }

    const { semanticColors, palette, fonts } = currentTheme;
    this.domElement.classList.toggle(styles.isDarkTheme, !!currentTheme.isInverted);

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--bodyDivider', semanticColors.bodyDivider || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
      this.domElement.style.setProperty('--inputBorder', semanticColors.inputBorder || null);
      this.domElement.style.setProperty('--inputBackground', semanticColors.inputBackground || null);
      this.domElement.style.setProperty('--inputText', semanticColors.inputText || null);
      this.domElement.style.setProperty('--buttonBackground', semanticColors.buttonBackground || null);
      this.domElement.style.setProperty('--buttonBackgroundHovered', semanticColors.buttonBackgroundHovered || null);
      this.domElement.style.setProperty('--buttonText', semanticColors.buttonText || null);
      this.domElement.style.setProperty('--primaryButtonBackground', semanticColors.primaryButtonBackground || null);
      this.domElement.style.setProperty('--primaryButtonBackgroundHovered', semanticColors.primaryButtonBackgroundHovered || null);
      this.domElement.style.setProperty('--primaryButtonText', semanticColors.primaryButtonText || null);
      this.domElement.style.setProperty('--disabledBackground', semanticColors.disabledBackground || null);
      this.domElement.style.setProperty('--disabledText', semanticColors.disabledText || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
      this.domElement.style.setProperty('--themeDark', palette.themeDark || null);
      this.domElement.style.setProperty('--themeLighterAlt', palette.themeLighterAlt || null);
      this.domElement.style.setProperty('--neutralLighterAlt', palette.neutralLighterAlt || null);
      this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || null);
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
      this.domElement.style.setProperty('--neutralTertiary', palette.neutralTertiary || null);
      this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || null);
      this.domElement.style.setProperty('--neutralPrimary', palette.neutralPrimary || null);
      this.domElement.style.setProperty('--white', palette.white || null);
    }

    if (fonts && fonts.medium) {
      this.domElement.style.setProperty('--bodyFontFamily', fonts.medium.fontFamily || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                // Bound to the existing "description" property from the scaffold; used as the list title.
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
