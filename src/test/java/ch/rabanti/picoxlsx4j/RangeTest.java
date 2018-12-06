package ch.rabanti.picoxlsx4j;

import ch.rabanti.picoxlsx4j.matchers.AddressMatchers;
import ch.rabanti.picoxlsx4j.matchers.RangeMatchers;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.allOf;

class RangeTest {

    @DisplayName("Should return a valid Range object according to the composite input data (Addresses)")
    @ParameterizedTest(name = "Start: {6}(C:{0},R:{1},T:{2}) and End: {7}(C:{3},R:{4},T:{5}) should lead to {10}")
    @CsvSource({
            "0,0,Default,0,0,Default,A1,B1,A1,B1,A1:B1",
            "0,0,Default,0,0,Default,$F4,$Z9,$F4,$Z9,$F4:$Z9"
    })
    public void compositeTest(int startCol, int startRow, Cell.AddressType startType, int endCol, int endRow, Cell.AddressType endType, String startAddr, String endAddr, String expectedStart, String expectedEnd, String expectedRange){
        Range r = buildRange(startCol,startRow,startType,endCol,endRow,endType,startAddr,endAddr);
        Address start = new Address(expectedStart);
        Address end = new Address(expectedEnd);
        assertThat(r,
                allOf(
                        RangeMatchers.hasStartAddress(start),
                        RangeMatchers.hasEndAddress(end),
                        RangeMatchers.hasRangeString(expectedRange)
                )
        );
    }

    @ParameterizedTest
    public void toStringTest() {
    }

    private static Range buildRange(int startCol, int startRow, Cell.AddressType startType, int endCol, int endRow, Cell.AddressType endType, String startAddr, String endAddr){
        Address start, end;
        if (startAddr == null){

            start = new Address(startCol, startRow, startType);
        }
        else {
            start = new Address(startAddr);
        }
        if (endAddr == null){
            end = new Address(endCol, endRow, endType);
        }
        else {
            end = new Address(endAddr);
        }
        return new Range(start, end);
    }
}
