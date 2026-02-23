import * as React from 'react';
import { Checkbox } from '@fluentui/react-components';

export interface IMultiOptionSetCheckboxProps {
    options: any[];
    selected: number[];
    disabled: boolean; // FIX 2: added disabled prop
    columns: number;
    rows: number;
    orderBy: string;
    direction: string;
    startAt: string;
    orientation: string;
    onChange: (selected: number[]) => void;
}

const MultiOptionSetCheckbox = ({ options, selected, disabled, columns, rows, orderBy, direction, startAt, orientation, onChange }: IMultiOptionSetCheckboxProps) => {
    const [splitOptions, setSplitOptions] = React.useState([] as any);
    const [calcColumns, setColumns] = React.useState(columns);
    const [calcRows, setRows] = React.useState(rows);

    React.useEffect(() => {
        let sortedOptions = [...options];
        if (orderBy === "optionsetvalue") {
            sortedOptions.sort((a, b) => a.Value - b.Value);
        } else if (orderBy === "alphabetical") {
            sortedOptions.sort((a, b) => a.Label.localeCompare(b.Label));
        }
        if (direction === "desc") {
            sortedOptions.reverse();
        }
        if (startAt && startAt !== "") {
            splitOptionsFunction(sortedOptions, startAt, orientation);
        } else {
            setSplitOptions(sortedOptions);
            // FIX 5: Reset columns/rows when startAt is not set, so the grid
            // layout is always consistent with the current props.
            setColumns(columns);
            setRows(rows);
        }
    }, [options, orderBy, direction, startAt, orientation, columns, rows]);

    const splitOptionsFunction = (options: any, startAt: any, orientation: any) => {
        options = options.sort((a: any, b: any) => a.Label.localeCompare(b.Label));
        const startCharacters = startAt.split(';').map((x: string) => x.toUpperCase());
        const optionsSplit: any[][] = Array.from({ length: startCharacters.length }, () => []);

        options.forEach((option: any) => {
            const firstLetter = option.Label.charAt(0).toUpperCase();
            const index = findIndexForLetter(firstLetter, startCharacters);
            optionsSplit[index].push(option);
        });

        const longestListLength = Math.max(...optionsSplit.map(list => list.length));
        const paddedOptions = optionsSplit.map(list => list.concat(new Array(longestListLength - list.length).fill(null))).flat();
        setSplitOptions(paddedOptions);

        const calcRows = orientation === "column" ? longestListLength : Math.ceil(paddedOptions.length / longestListLength);
        const calcColumns = orientation === "column" ? Math.ceil(paddedOptions.length / longestListLength) : longestListLength;
        setColumns(calcColumns);
        setRows(calcRows);
    };

    const findIndexForLetter = (letter: any, startCharacters: any) => {
        for (let i = 0; i < startCharacters.length; i++) {
            if (startCharacters[i] <= letter && letter < startCharacters[i + 1]) {
                return i;
            }
        }
        return startCharacters.length - 1;
    };

    const handleCheckboxChange = (option: any) => {
        // FIX 2: Do nothing if control is disabled
        if (disabled) return;

        const updatedSelected = [...selected];
        // FIX 1 & 4: Coerce to Number for comparison so that string "123" and
        // number 123 are treated as the same value. Dataverse can return option
        // Values as strings in some PCF runtime versions, causing indexOf / includes
        // to miss matches and leaving checkboxes unchecked or producing duplicates.
        const optionValue = Number(option.Value);
        const index = updatedSelected.findIndex(v => Number(v) === optionValue);
        if (index > -1) {
            updatedSelected.splice(index, 1);
        } else {
            updatedSelected.push(optionValue);
        }
        onChange(updatedSelected);
    };

    // FIX 1: Normalise selected values to Numbers once at render time so that
    // the checked state comparison is always type-safe.
    const normalizedSelected = React.useMemo(
        () => (selected ?? []).map(v => Number(v)),
        [selected]
    );

    return (
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${calcColumns}, auto)`, gridTemplateRows: `repeat(${calcRows}, auto)`, gridAutoFlow: orientation === "column" ? 'column' : 'row' }}>
            {splitOptions.map((option: any, index: number) => (
                <div key={index}>
                    {option && (
                        <Checkbox
                            label={option.Label}
                            // FIX 1: Use normalizedSelected (Numbers) and coerce option.Value
                            // to Number so the comparison is always type-safe.
                            checked={normalizedSelected.includes(Number(option.Value))}
                            // FIX 2: Pass disabled to the Fluent UI Checkbox so the control
                            // is correctly read-only when the form is deactivated.
                            disabled={disabled}
                            onChange={() => handleCheckboxChange(option)}
                        />
                    )}
                </div>
            ))}
        </div>
    );
};

export default MultiOptionSetCheckbox;