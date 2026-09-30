export function isJsonObject(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
export function isJsonValue(value) {
    if (value === null || typeof value === "string" || typeof value === "boolean")
        return true;
    if (typeof value === "number")
        return Number.isFinite(value);
    if (Array.isArray(value))
        return value.every(isJsonValue);
    if (!isJsonObject(value))
        return false;
    return Object.values(value).every(isJsonValue);
}
export function cloneJson(value) {
    if (Array.isArray(value)) {
        return value.map((entry) => cloneJson(entry));
    }
    if (isJsonObject(value)) {
        return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, cloneJson(entry)]));
    }
    return value;
}
export function jsonEquals(left, right) {
    if (left === right)
        return true;
    if (Array.isArray(left) || Array.isArray(right)) {
        return (Array.isArray(left) &&
            Array.isArray(right) &&
            left.length === right.length &&
            left.every((entry, index) => {
                const other = right[index];
                return other !== undefined && jsonEquals(entry, other);
            }));
    }
    if (isJsonObject(left) || isJsonObject(right)) {
        if (!isJsonObject(left) || !isJsonObject(right))
            return false;
        const leftKeys = Object.keys(left).sort();
        const rightKeys = Object.keys(right).sort();
        return (leftKeys.length === rightKeys.length &&
            leftKeys.every((key, index) => {
                const leftValue = left[key];
                const rightValue = right[key];
                return (key === rightKeys[index] &&
                    leftValue !== undefined &&
                    rightValue !== undefined &&
                    jsonEquals(leftValue, rightValue));
            }));
    }
    return false;
}
