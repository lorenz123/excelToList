import lombok.Data;

import java.math.BigInteger;
import java.util.List;

@Data
public class Employee {
    private String employeeId;
    private String fullName;
    private String nickName;
    private String newNickname;
    private Double uuid;
    private String depositAddress;
    private List<EmployeeDetails> employeeDetails;
}
