#include <windows.h> 


int test() { 
        MessageBoxW(NULL, L"The message", L"The caption", MB_OK);

        return 0; 
} 

int main() {
	return test();
}
