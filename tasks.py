from invoke import task

@task(default = True)
def test(c):
    c.run('python main.py')

# @task
# def test(c):
#     c.run('python -m unittest discover -s tests -p "test_*.py"')

# @task
# def flake8(c):
#     c.run('flake8 xlcrf')
    
