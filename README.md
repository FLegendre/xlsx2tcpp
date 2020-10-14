# xlsx2tcpp
Convert a Microsoft xlsx worksheet into a table that is read efficiently in C++ thanks to many threads.

This is an header only library : put the `xlsx2tcpp.hpp` file in a appropriate place and use it with 3 steps.

First step: generate the `test-sheet1.hpp` file from the `test.xlsx` workbook using the C++ program:
```C++
#include <xlsx2tcpp.hpp>

int
main()
{

    xlsx2tcpp::init("test.xlsx");

    return 0;
}
```

Second step: build the zipped chunks of the `sheet1` sheet using the C++ program:
```C++
#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

    xlsx2tcpp::build<test_sheet1>("test.xlsx");

    return 0;
}
```

Third step: use the zipped chunks of the `sheet1` sheet using the C++ program:
```C++
#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

    auto const table{ xlsx2tcpp::read<test_sheet1>() };

    for (auto const& row : table)
        std::cout << row.b << '\n';

    return 0;
}
```
