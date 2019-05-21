package ch.rabanti.picoxlsx4j;

import ch.rabanti.picoxlsx4j.matchers.AddressMatchers;
import ch.rabanti.picoxlsx4j.matchers.RangeMatchers;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.allOf;
import static org.hamcrest.Matchers.is;

class RangeTest {

    @DisplayName("Should return a valid Range object according to the composite input data (Addresses)")
    @ParameterizedTest(name = "Start: {0} and End: {1} should lead to {4} ({2}:{3})")
    @CsvSource({
            "A1,B1,A1,B1,A1:B1",
            "$F4,$Z9,$F4,$Z9,$F4:$Z9",
            "$F$4,$Z$9,$F$4,$Z$9,$F$4:$Z$9",
            "$X12,R555,$X12,R555,$X12:R555",
            "S18,$Z29,S18,$Z29,S18:$Z29"
    })
    public void compositeTest(String startAddr, String endAddr, String expectedStart, String expectedEnd, String expectedRange){
        Range r = buildRange(startAddr,endAddr);
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

    @DisplayName("Should return a valid Range as String")
    @ParameterizedTest(name = "Start: {0} and End: {1} should lead to {2}")
    @CsvSource({
            "A1,B1,A1:B1",
            "$C1100,X200,$C1100:X200",
            "$D$11,Z200,$D$11:Z200",
            "S22,$V50, S22:$V50",
            "L8,$X$9, L8:$X$9"
    })
    public void toStringTest(String startAddr, String endAddr, String expectedString) {
        Range r = buildRange(startAddr, endAddr);
        assertThat(r.toString(), is(expectedString));
    }

    public void invalidTest(){

    }

    private static Range buildRange(String startAddr, String endAddr){
        Address start, end;
        start = new Address(startAddr);
        end = new Address(endAddr);
        return new Range(start, end);
    }
}
