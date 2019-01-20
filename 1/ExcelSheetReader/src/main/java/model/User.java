package model;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;


@Getter
@Setter
@EqualsAndHashCode
@ToString
public class User {
    Integer indexId;
    String firstName;
    String lastName;
    String gender;
    String country;
    Integer age;
    Date date;
    Integer id;

}
