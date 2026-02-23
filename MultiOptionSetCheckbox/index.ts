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

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        // Only load options here (they don't change). Do NOT cache _selected here —
        // the field value may not be loaded yet at init time.
        this._options = context.parameters.MultiSelectColumn.attributes?.Options || [];
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        // FIX 1: Always read the current selected values from context in updateView,
        // not from a value cached during init. This ensures that when the form loads
        // and populates the field, the control reflects the correct selections.
        const rawValue = context.parameters.MultiSelectColumn.raw;
        this._selected = Array.isArray(rawValue) ? rawValue : [];

        // FIX 2: Read the disabled state from context so the control correctly
        // becomes read-only when the form is deactivated or the field is locked.
        const isDisabled =
            context.mode.isControlDisabled ||
            context.parameters.MultiSelectColumn.security?.readable === false;

        ReactDOM.render(
            React.createElement(MultiOptionSetCheckbox, {
                options: this._options,
                selected: this._selected,
                disabled: isDisabled,
                columns: context.parameters.Columns.raw || 1,
                rows: context.parameters.Rows.raw ?? Math.ceil(this._options.length / (context.parameters.Columns.raw || 1)),
                orderBy: context.parameters.OrderBy.raw,
                direction: context.parameters.Direction.raw,
                startAt: context.parameters.Startat.raw,
                orientation: context.parameters.Orientation.raw,
                onChange: (selected: number[]) => {
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