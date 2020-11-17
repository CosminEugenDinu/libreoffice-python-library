
import sys

def Messages(label=''):
    """
    Call inside a function like:
    add_msg, get_msgs = Messages('warning')

    Add debug messages like this:
    add_msg('first mess')
    add_msg('sec meg')
    add_msg('first_mess')

    Retieve messages like this:
    messages = get_msgs()
    print(messages)
    """

    _who_called_me = sys._getframe(1).f_code.co_name
    _msgs_count = {}

    def add_msg(msg):
        nonlocal _msgs_count

        if label:
            _msgs_count['label'] = label
        if type(msg) is not str:
            raise Exception('msg must be of type str')

        if _msgs_count.get(msg) is not None:
            _msgs_count[msg] += 1
        else:
            _msgs_count[msg] = 1

        return True
    
    def get_msgs():

        if len(_msgs_count) > 0:

            return {
            _who_called_me: {
                **_msgs_count
                }
            }

    return add_msg, get_msgs