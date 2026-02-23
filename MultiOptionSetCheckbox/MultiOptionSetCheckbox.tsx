import * as React from 'react';
import { Checkbox } from '@fluentui/react-components';

export interface IMultiOptionSetCheckboxProps {
    options: any[];
    selected: number[];
    disabled: boolean;
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

    // Normalise selected values to Numbers once per render so that the checked
    // state comparison is always type-safe. The PCF runtime can surface option
    // Values as strings in some versions, causing strict equality checks to fail.
    const normalizedSelected = React.useMemo(
        () => (selected ?? []).map(v => Number(v)),
        [selected]
    );

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
            // Reset grid dimensions to prop values when startAt is not in use,
            // preventing stale dimensions left over from a previous splitAt calculation.
            setColumns(columns);
            setRows(rows);
        }
    }, [options, orderBy, direction, startAt, orientation, columns, rows]);

    const splitOptionsFunction = (options: any, startAt: any, orientation: any) => {
        // Sort options alphabetically
        options = options.sort((a: any, b: any) => a.Label.localeCompare(b.Label));

        // Parse startAt string and convert to uppercase
        const startCharacters = startAt.split(';').map((x: string) => x.toUpperCase());

        // Create a new array to store the split options
        const optionsSplit: any[][] = Array.from({ length: startCharacters.length }, () => []);

        // Create sublists for each starting character
        options.forEach((option: any) => {
            const firstLetter = option.Label.charAt(0).toUpperCase();
            const index = findIndexForLetter(firstLetter, startCharacters);
            optionsSplit[index].push(option);
        });

        // Pad sublists with null values to make them equal in length
        const longestListLength = Math.max(...optionsSplit.map(list => list.length));
        const paddedOptions = optionsSplit.map(list => list.concat(new Array(longestListLength - list.length).fill(null))).flat();
        setSplitOptions(paddedOptions);

        // Calculate rows and columns based on orientation
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
        // Belt-and-suspenders guard — isDisabled is also enforced in index.ts onChange.
        if (disabled) return;

        const updatedSelected = [...selected];
        // Coerce both sides to Number to avoid string/number mismatch when finding
        // and removing an existing selection, which could otherwise cause duplicates.
        const optionValue = Number(option.Value);
        const index = updatedSelected.findIndex(v => Number(v) === optionValue);
        if (index > -1) {
            updatedSelected.splice(index, 1);
        } else {
            updatedSelected.push(optionValue);
        }
        onChange(updatedSelected);
    };

    return (
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${calcColumns}, auto)`, gridTemplateRows: `repeat(${calcRows}, auto)`, gridAutoFlow: orientation === "column" ? 'column' : 'row' }}>
            {splitOptions.map((option: any, index: number) => (
                <div key={index}>
                    {option && (
                        <Checkbox
                            label={option.Label}
                            // Use normalizedSelected and coerce option.Value to Number
                            // so the checked comparison is always type-safe.
                            checked={normalizedSelected.includes(Number(option.Value))}
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