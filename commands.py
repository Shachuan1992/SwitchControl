def write_command(instance,command):
    finish = '\r'
    instance.write(command.encode('ascii')+finish.encode('ascii'))