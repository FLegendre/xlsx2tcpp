# xlsx2tcpp
Convert a Microsoft xlsx worksheet into a table that is read efficiently in C++ thanks to many threads.

Presume that the worksheet is
|  a  |  b  |  c  |
| --- | --- | --- |
|  1  | m   |  1. |
|  2  | mm  |  2. |
|  1  | m   |  3. |

The library will generate a C++ struct with 3 data members, named `a`, `b`, and `c`, of type `int64_t`, `std::array<char, 2>`, and `double`. In a second step, the library generate zipped chunks of the original sheet, with respect to the number of cores of the processor.

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

This library depends on the libzip library: https://libzip.org/, the zlib library: https://zlib.net/, and the fd-xlsx-read library: https://github.com/FLegendre/fd-read-xlsx.
