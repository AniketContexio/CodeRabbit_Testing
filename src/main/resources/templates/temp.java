import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

public class HelloTest {

    Hello hello = new Hello();

    @Test
    public void testGreetWithName() {
        String result = hello.greet("Aniket");
        assertEquals("Hello, Aniket!", result);
    }

    @Test
    public void testGreetWithEmptyString() {
        String result = hello.greet("");
        assertEquals("Hello, !", result);
    }

    @Test
    public void testGreetWithNull() {
        String result = hello.greet(null);
        assertEquals("Hello, null!", result);
    }
}
