import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import MultiOptionSetCheckbox, { IMultiOptionSetCheckboxProps } from './MultiOptionSetCheckbox'; // Path to your React component file

export class MultiOptionsetCheckbox implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _options: any[] = [];
    private _selected: number[] = [];

    private _notifyOutputChanged: () => void;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        this._options = this._context.parameters.MultiSelectColumn.attributes?.Options || [];
        this._selected = this._context.parameters.MultiSelectColumn.raw || [];
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Always use the latest context passed by the framework
        this._context = context;

        // Rehydrate options + selected values every time (fixes inconsistent initial load)
        this._options = context.parameters.MultiSelectColumn.attributes?.Options || [];
        this._selected = [...(context.parameters.MultiSelectColumn.raw || [])];

        // Respect read-only/disabled state
        const isDisabled = context.mode.isControlDisabled;

        ReactDOM.render(
            React.createElement(MultiOptionSetCheckbox, {
                options: this._options,
                selected: this._selected,
                disabled: isDisabled,
                columns: context.parameters.Columns.raw || 1,
                rows: context.parameters.Rows.raw || Math.ceil(this._options.length / (context.parameters.Columns.raw || 1)),
                orderBy: context.parameters.OrderBy.raw,
                direction: context.parameters.Direction.raw,
                startAt: context.parameters.Startat.raw,
                orientation: context.parameters.Orientation.raw,
                onChange: (selected) => {
                    // Block updates when form/control is read-only
                    if (isDisabled) return;

                    this._selected = selected;
                    this._notifyOutputChanged();
                }
            } as IMultiOptionSetCheckboxProps),
            this._container
        );
    }


    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs
    {
        return { MultiSelectColumn: this._selected };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}