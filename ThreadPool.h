/****************************** Module Header ******************************\
Module Name:  ThreadPool.h
Project:      CppShellExtContextMenuHandler

Creates a pool of threads that take tasks. Intended for the maximum number of 
threads available for hardware or any other number of threads.

\***************************************************************************/

#pragma once

#ifndef THREADPOOL_H
#define THREADPOOL_H

#include <iostream>
#include <thread>
#include <mutex>
#include <condition_variable>
#include <queue>
#include <functional>
#include <chrono>

class ThreadPool
{
public:
    //! Haponov - create as many threads as needed
    ThreadPool(int threads);

    //!Haponov - hand tasks to threads of pool
    void doJob(std::function <void(void)> func);

    ~ThreadPool();

protected:
    //!Haponov - manage threads/tasks in pool:
    //           move tasks in container, actully run the tasks
    void threadEntry(int i);

    std::mutex lock_;
    std::condition_variable condVar_;
    bool shutdown_;

    //!Haponov - contain tasks to do
    std::queue <std::function <void(void)>> jobs_;
    //!Haponov - threads for doing tasks
    std::vector <std::thread> threads_;
};

#endif // THREADPOOL_H

