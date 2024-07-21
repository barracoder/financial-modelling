import { add, clock } from "../../functions/functions"; 



beforeEach(() => {
    jest.useFakeTimers();
    jest.spyOn(global, "setInterval");
    jest.spyOn(global, "clearInterval");
});

afterEach(() => {
    jest.clearAllMocks();
});

describe("add function", () => {
    it("correctly adds two numbers", () => {
        expect(add(1, 2)).toBe(3);
        expect(add(-1, -2)).toBe(-3);
        expect(add(0, 0)).toBe(0);
    });
});

describe("clock function", () => {
    jest.useFakeTimers();
    const mockSetResult = jest.fn();
    const mockOnCanceled = jest.fn();

    const mockInvocation: CustomFunctions.StreamingInvocation<string> = {
        setResult: mockSetResult,
        onCanceled: mockOnCanceled,
    };

    it("sets the result with the current time every second", () => {
        clock(mockInvocation);
        expect(setInterval).toHaveBeenLastCalledWith(expect.any(Function), 1000);

        jest.advanceTimersByTime(1000); // Simulate 1 second passing
        expect(mockSetResult).toHaveBeenCalledTimes(1);

        jest.advanceTimersByTime(3000); // Simulate 3 more seconds passing
        expect(mockSetResult).toHaveBeenCalledTimes(4);
    });

    it("clears the interval when canceled", () => {
        clock(mockInvocation);
        expect(setInterval).toHaveBeenCalledTimes(1); // 

        mockInvocation.onCanceled();
        jest.advanceTimersByTime(1000); // Simulate 1 second passing
        expect(mockSetResult).toHaveBeenCalledTimes(0); // No additional calls after cancelation
    });
});
