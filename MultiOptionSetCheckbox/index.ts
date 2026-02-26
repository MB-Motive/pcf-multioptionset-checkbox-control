import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import MultiOptionSetCheckbox, { IMultiOptionSetCheckboxProps } from './MultiOptionSetCheckbox';

export class MultiOptionsetCheckbox implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _options: any[] = [];
    private _selected: number[] = [];
    private _notifyOutputChanged: () => void;

    constructor() { }

    private arraysEqual(a: number[], b: number[]): boolean {
        if (a === b) return true;
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (a[i] !== b[i]) return false;
        }
        return true;
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        // Load options in init — they are available immediately and don't change.
        // Do NOT cache _selected here; field data may not be loaded yet at init time.
        this._options = context.parameters.MultiSelectColumn.attributes?.Options || [];
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        // Rehydrate options every updateView cycle.
        this._options = context.parameters.MultiSelectColumn.attributes?.Options || [];

        // Rehydrate selected values with a precise undefined/null distinction:
        //   undefined => field data not loaded yet — do NOT overwrite current selection.
        //   null      => field loaded and is empty — treat as empty array.
        const rawSelected = context.parameters.MultiSelectColumn.raw;
        if (rawSelected !== undefined) {
            const nextSelected = rawSelected ?? [];
            if (!this.arraysEqual(nextSelected, this._selected)) {
                this._selected = nextSelected;
            }
        }

        // Respect form disabled state and field-level security.
        // Uses .editable (not .readable) — the correct property for write access.
        const isDisabled =
            context.mode.isControlDisabled ||
            !context.parameters.MultiSelectColumn.security?.editable;

        ReactDOM.render(
            React.createElement(MultiOptionSetCheckbox, {
                options: this._options,
                selected: this._selected,
                disabled: isDisabled,
                columns: context.parameters.Columns.raw || 1,
                rows: context.parameters.Rows.raw ||
                    Math.ceil(this._options.length / (context.parameters.Columns.raw || 1)),
                orderBy: context.parameters.OrderBy.raw,
                direction: context.parameters.Direction.raw,
                startAt: context.parameters.Startat.raw,
                orientation: context.parameters.Orientation.raw,
                onChange: (selected: number[]) => {
                    // Belt-and-suspenders guard in addition to the TSX-level guard.
                    if (isDisabled) return;
                    this._selected = selected;
                    this._notifyOutputChanged();
                }
            } as IMultiOptionSetCheckboxProps),
            this._container
        );
    }

    public getOutputs(): IOutputs {
        return { MultiSelectColumn: this._selected };
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}